# jevOutlook — Architecture (miroir français)

> Source de vérité : `ARCHITECTURE_EN.md`. Garder les deux fichiers synchronisés.

## 1. Objectif

Portage du classifieur jevMail (Gmail / Apps Script) vers l'écosystème Microsoft :
classer les messages Outlook avec le modèle de décision Jev (TypeSafe, via OpenRouter),
appliquer une catégorie Outlook par message, archiver optionnellement le courrier
jetable à haute confiance, sous un budget strict et avec un traitement reprenable
et checkpointé.

## 2. Pile technique

| Sujet | Choix | Pourquoi |
| --- | --- | --- |
| Langage / runtime | C# 13, .NET 10 (`net10.0`), application console | Langage de la plateforme Microsoft ; une CLI est l'équivalent naturel d'un Apps Script « coller et lancer » |
| API boîte mail | Microsoft Graph v1.0 REST via `HttpClient` | Contrôle total de `$batch`, des en-têtes `Prefer` et des filtres OData ; pas de dérive de version du SDK |
| Connexion | `Azure.Identity` (`InteractiveBrowserCredential`, `DeviceCodeCredential`) | Permissions déléguées, cache de jetons MSAL persistant, renouvellement silencieux via `AuthenticationRecord` |
| Modèle | Jev `~typesafe/jev-latest` via OpenRouter Decisions (`/api/alpha/decisions`) ; TypeSafe direct optionnel (`/v1/systemone`) | Même contrat requête/réponse ; OpenRouter ajoute `usage.cost` |
| État | Fichiers JSON dans `~/.jevoutlook/` (0600) | Équivalent des User Properties d'Apps Script |
| Tests | xUnit | Logique pure : règles, payload, parsing du contrat, compaction du corps, filtres, options |

## 3. Carte des composants

```
Cli/Program.cs ──► Triage/TriageEngine ──► Graph/GraphMailClient ──► Microsoft Graph
      │                    │                       ▲
      │                    ├──► Jev/JevClient ─────┼──► OpenRouter / TypeSafe
      │                    │                       │
      │                    └──► Storage/JobStore, RuleStore      Graph/GraphAuth (Azure.Identity)
      └──► Storage/AppConfig, Rules/Playbooks, Rules/RuleValidator
```

| Module | Responsabilité |
| --- | --- |
| `AppConstants` | Tous les réglages (fenêtres de concurrence, délais de retry, seuils, nom du marqueur, défauts) |
| `Rules/*` | Record `LabelRule`, 7 playbooks (identiques à jevMail), validation (≤12 règles, pas de `,`/`;`, pas de marqueur réservé, régénération des ids) |
| `Storage/*` | Chemins, écritures JSON atomiques en 0600, `AppConfig`, `RuleStore` (règles sauvegardées + règles figées du job), `JobStore` |
| `Graph/GraphAuth` | Construction des credentials, connexion interactive, persistance du record, obtention des jetons |
| `Graph/GraphMailClient` | Listage par curseur, lectures `$batch`, catégories maîtres, PATCH des catégories, déplacement vers Archive, gestion du throttling |
| `Graph/MessageMetadata` | Normalise un message Graph en objet `email` plat envoyé à Jev ; HTML→texte ; compaction tête/queue |
| `Jev/JevPayloadBuilder` | Construit la question Choice (options `L0..Ln`), garde-fou de taille (29 000 octets), estimation de coût |
| `Jev/JevClient` | Envoie une requête Decisions ; validation stricte du contrat ; source du coût |
| `Triage/TriageEngine` | Cycle de vie de session, lots, deux vagues, concurrence adaptative, budget, disjoncteur, écritures groupées |
| `Cli/*` | Commandes, analyse des arguments, rendu terminal |

## 4. Correspondance Gmail → Outlook

| Concept | Gmail (jevMail) | Outlook (jevOutlook) |
| --- | --- | --- |
| Label | Label Gmail (créé via `labels.create`) | Catégorie : `POST /me/outlook/masterCategories` `{displayName, color: presetN}` puis `PATCH /me/messages/{id}` `{categories:[…]}` (remplacement complet → les catégories existantes sont fusionnées, jamais supprimées) |
| Archivage | `messages.batchModify` retire `INBOX` | `POST /me/messages/{id}/move` `{destinationId:"archive"}` |
| Marqueur | label `jev-triaged` | catégorie `jev-triaged` (couleur `none`) |
| Périmètre | `in:inbox` / `-in:trash -in:spam` | `/me/mailFolders/inbox/messages` / `/me/messages` avec exclusion côté client de Courrier indésirable, Éléments supprimés, Brouillons (ids de dossiers résolus une fois) |
| Non lus | `is:unread` | `isRead eq false` |
| Sauter le déjà traité | `-label:jev-triaged` dans la requête | côté client : ignorer les refs dont `categories` contient le marqueur (Graph n'a pas de `not categories/any` fiable) |
| Pagination | `pageToken` (preview) / la requête se réduit (live) | curseur `receivedDateTime`, voir §5 |
| Lecture métadonnées | `messages.get format=metadata` + en-têtes choisis, batch HTTP multipart | `$batch` de `GET /me/messages/{id}?$select=…,internetMessageHeaders` (20 par batch) |
| Lecture complète | `messages.get format=full` + parcours MIME | `$batch` de `GET /me/messages/{id}?$select=id,body` avec `Prefer: outlook.body-content-type="text"` |
| Estimation | `resultSizeEstimate` | Inbox : `unreadItemCount` / `totalItemCount` ; tout : `$count=true` (peut être indisponible → inconnu) |
| Limites de débit | 429 + Retry-After, fenêtre 25→50 | 429/503 + Retry-After ; ≤2 appels `$batch` en vol (`SemaphoreSlim`, permis conservé jusqu'à la lecture du corps — chaque requête interne compte dans les ~4 concurrentes par app × boîte), fenêtre 20→40 messages |

## 5. Curseur de listage

Graph exige que les propriétés de `$orderby` apparaissent en premier dans `$filter`. Le moteur
liste donc du plus récent au plus ancien avec :

```
$filter=receivedDateTime le {curseur}[ and isRead eq false]&$orderby=receivedDateTime desc&$top=N
```

- Le curseur n'avance que sur les messages **examinés** (collectés ou ignorés) ; une page peut
  donc contenir plus de candidats qu'un lot n'en a besoin sans en perdre aucun.
- Les ids partageant l'horodatage du curseur sont conservés dans `CursorBoundaryIds` et exclus
  côté client : les ex æquo ne sont ni répétés ni perdus. Si une page pleine ne contient que des
  ids de frontière (plus d'horodatages identiques qu'une page n'en tient), la requête suivante passe
  en `lt` (exclusif) : rien n'est perdu ni dupliqué entre-temps ; seuls d'autres messages portant
  exactement cet horodatage seraient laissés de côté. Les ids collectés sont dédoublonnés par sécurité.
- Le curseur étant temporel, archiver des messages hors de la boîte de réception en cours de run ne
  décale jamais les pages, et les runs live et preview utilisent le même mécanisme.
- Les messages reçus après le début de la session (curseur = début + 5 min) n'en font pas partie.

## 6. Algorithme de traitement (un lot)

1. **Remplir les en-attente** (≤ 50) depuis le curseur ; terminer quand épuisé ou limite atteinte.
2. **Lecture des métadonnées** (fenêtre `$batch` adaptative). Les éléments portant déjà le marqueur
   sans décision finale sont retirés (traités ailleurs). 404 → ignoré sans risque.
3. **Vague 1 — métadonnées** : une requête Choice indépendante par élément sans résultat.
   Les réponses réussies sont checkpointées avant le traitement des échecs.
4. **Décision** : confiance ≥ seuil métadonnées (et, en mode archive pour une catégorie éligible,
   ≥ seuil d'archivage) → finale ; sinon l'élément passe en vague 2.
5. **Lecture complète** + extraction du texte. Corps vide → repli sûr `review` (non archivable)
   si une telle règle existe, sinon ignoré.
6. **Vague 2 — corps complet** : requêtes indépendantes ; réponses checkpointées.
7. **Préfixe prêt** : les premiers éléments en attente ayant une décision finale sont écrits en deux
   passes groupées idempotentes — les `categories` courantes de chaque message sont **relues au même
   instant** (un PATCH remplace toute la collection et l'instantané du lot peut dater de plusieurs
   minutes), puis `PATCH categories` (courant ∪ {catégorie, marqueur}), puis `move` pour les décisions
   d'archivage. Un échec Graph transitoire met la session en pause avec toutes les décisions
   sauvegardées. Au replay, un élément en attente qui a déjà une décision finale mais répond 404 (son id
   a changé parce que le déplacement vers Archive a réussi avant le checkpoint) est compté comme traité,
   pas comme ignoré.
8. **Compteurs** (traités, métadonnées seules, complets, archivés, par catégorie), tableau de
   résultats, en-attente réduits, job sauvegardé. Blocage budget → statut `budget`.

### Envoi des vagues (`DispatchWaveAsync`)

- Réservation du budget avant envoi (`SpentUsd + réserve + estimation ≤ MaxSpendUsd`), remplacée
  ensuite par le coût réel (`reported` > `input_tokens` > `estimated`).
- Les requêtes partent en parallèle par fenêtres de `JevConcurrency` (25 → 50, divisée par deux en cas d'échec).
- Échecs transitoires de portée session (429/5xx/transport) : si toute la fenêtre a échoué avec le
  même code, **sonder une requête** avant de rejouer le reste ; sinon rejouer une fois les échecs.
- Échecs de contrat de portée message (JSON invalide, option inconnue, probabilités incorrectes) :
  rejoués une fois, bornés par le disjoncteur (3 consécutifs → pause).

### Politique d'échec (fail closed)

| Situation | Effet |
| --- | --- |
| Fournisseur 401/402/403/4xx | Erreur / pause de session, aucun changement Outlook |
| Fournisseur 429/5xx après retry | Pause `provider-temporary`, reprenable. Un 429 n'est pas facturé (réservation libérée) ; timeouts/5xx gardent l'estimation prudente |
| Réponse modèle invalide ×1 | Retry une fois ; puis message ignoré, `ModelResponseSkips++` |
| 3 réponses invalides consécutives | Pause `jev-response-circuit-breaker` |
| Graph 429/5xx après retry | Pause `graph-temporary` (lectures) ou après écriture groupée (`graph-temporary`) |
| Graph 401/403 | Erreur, demande de reconnexion |
| Message 404 | Ignoré sans risque (lectures) ; considéré comme fait (écritures) ; compté comme traité si l'élément porte déjà une décision finale (replay après déplacement vers Archive) |
| 40 pages de listage sans nouveau candidat | Pause `scan-guard` (tout porte déjà le marqueur) ; `continue` poursuit le balayage |
| Ctrl+C pendant une vague | Les réservations des requêtes sans réponse sont libérées ; elles sont renvoyées à la reprise |
| Budget dépassé | Statut `budget` ; `continue --max-spend` pour relever |
| Ctrl+C | Pause `user-stop` ; `continue` reprend (et réarme le disjoncteur) |

## 7. Persistance

| Fichier | Contenu |
| --- | --- |
| `config.json` | client id, tenant, fournisseur, surcharges endpoint/modèle, clé API optionnelle, drapeau device-code |
| `auth-record.json` | `AuthenticationRecord` MSAL (aucun secret ; les jetons vivent dans le cache de l'OS) |
| `rules.json` | règles sauvegardées |
| `job.json` | session courante : options, curseur, compteurs, dépense, éléments en attente avec décisions checkpointées |
| `job-rules.json` | règles figées pour la session courante |

Tous les fichiers sont écrits atomiquement (`.tmp` + déplacement) en mode 0600 ; le répertoire est en 0700.

## 8. Notes de sécurité

- Permissions déléguées uniquement (`Mail.ReadWrite`, `MailboxSettings.ReadWrite`, `User.Read`) ;
  l'application n'envoie jamais de courrier.
- Les corps d'erreur du fournisseur ne sont jamais réaffichés (ils peuvent contenir des données de requête).
- Le contenu des e-mails est transmis au modèle comme donnée non fiable ; les instructions
  demandent au modèle de ne jamais suivre des consignes trouvées dans l'e-mail.
- La clé API est lue d'abord dans l'environnement ; son stockage sur disque est optionnel.

## 9. État de vérification (2026-09-22)

- `dotnet build` : 0 avertissement, 0 erreur. `dotnet test` : 48 tests verts.
- Passe de revue de code indépendante (agent séparé) + vérification documentaire des faits Graph/Azure.Identity/TypeSafe ;
  constats Major corrigés (comptage des 404 au replay, garde de balayage, relecture des catégories avant écriture,
  fuite de réservation budgétaire à l'annulation, échecs Graph transitoires → pause au lieu d'erreur).
- Vérification réelle : chemin d'erreur OpenRouter exercé avec une clé invalide (HTTP 401 → message propre).
- Pas encore exercé contre une vraie boîte mail (nécessite une inscription d'application Entra) :
  filtre de listage, formes `$batch` et écritures catégories/déplacement validés uniquement contre
  la documentation Graph.
