# jevOutlook — Architecture (miroir français)

> Source de vérité : `ARCHITECTURE_EN.md`. Garder les deux versions synchronisées.

## 1. Objectif

Portage du classifieur jevMail (Gmail / Apps Script) en **client local multi-boîtes** :
classer les messages de plusieurs boîtes — Microsoft 365 / Outlook.com (Graph), Gmail
(IMAP + mot de passe d'application) et tout serveur IMAP — avec le modèle de décision Jev
(TypeSafe, via OpenRouter), poser un libellé par message (catégorie Outlook, label Gmail
ou mot-clé IMAP), archiver éventuellement le courrier jetable à haute confiance, sous un
budget strict et avec un traitement reprenable et checkpointé. Aucun backend hébergé,
aucune configuration Azure au-delà d'une inscription d'application partagée par les
boîtes Microsoft.

## 2. Pile technique

| Sujet | Choix | Pourquoi |
| --- | --- | --- |
| Langage / runtime | C# 13, .NET 10 (`net10.0`), un exécutable : tableau de bord web local + CLI | Langage de la plateforme Microsoft ; le tableau de bord reflète la web app Apps Script, la CLI sert au scripting |
| Boîtes Microsoft | Microsoft Graph v1.0 REST via `HttpClient` | Contrôle total de `$batch`, des en-têtes `Prefer` et des filtres OData ; Exchange Online n'a plus d'IMAP par mot de passe, Graph + connexion Microsoft est la seule voie |
| Connexion Microsoft | `Azure.Identity` (`DeviceCodeCredential`, `InteractiveBrowserCredential`) | Permissions déléguées, cache MSAL persistant, un `AuthenticationRecord` par compte ; une seule inscription Entra partagée par toutes les boîtes Microsoft |
| Boîtes Gmail / IMAP | MailKit (`ImapClient`) | Client IMAP mature avec extensions Gmail (`X-GM-LABELS`), mots-clés, dossiers special-use |
| Mots de passe | Trousseau macOS via `/usr/bin/security` (service `jevoutlook`, compte = id du compte) ; `secrets.json` 0600 ailleurs | Rien de secret à côté de la configuration ; aucune dépendance supplémentaire |
| Modèle | Jev `~typesafe/jev-latest` via OpenRouter Decisions (`/api/alpha/decisions`) ; TypeSafe direct optionnel (`/v1/systemone`) | Même contrat requête/réponse ; OpenRouter ajoute `usage.cost` |
| État | Fichiers JSON dans `~/.jevoutlook/` (0600), sous-répertoires par compte | Équivalent des User Properties d'Apps Script |
| Tests | xUnit | Logique pure : règles, payload, contrat, compaction du corps, filtres, pagination par curseur, helpers IMAP |

## 3. Carte des composants

```
Cli/Program.cs ─┐                                   ┌─► Graph/GraphMailClient ──► Microsoft Graph
Web/UiServer  ──┼─► Triage/TriageEngine ─► Mail/IMailbox ┤
                │          │                        └─► Imap/ImapMailbox (MailKit) ──► imap.gmail.com / tout IMAP
                │          ├─► Jev/JevClient ──► OpenRouter / TypeSafe
                │          └─► Storage/JobStore (par compte)
                └─► Mail/MailboxFactory ─► Storage/AccountStore, SecretStore, Graph/GraphAuth
```

| Module | Responsabilité |
| --- | --- |
| `AppConstants` | Tous les réglages (fenêtres de concurrence, délais de retry, seuils, valeurs par défaut) |
| `Rules/*` | Record `LabelRule`, 7 playbooks (identiques à jevMail), validation (≤ 12 règles, pas de `,`/`;`, régénération des ids) |
| `Mail/IMailbox` | La surface exacte dont le moteur a besoin : identité, capacités, estimation, listing par curseur, lectures groupées (`ReadPart`), `EnsureLabels`, `ApplyLabels`, `Archive`, nettoyage. Ids et curseurs sont des chaînes opaques du fournisseur ; `MailboxException(transient)` décide entre pause et erreur |
| `Mail/MailboxFactory` | Ouvre le bon fournisseur pour un `MailAccount` (Graph : record de connexion ; IMAP : mot de passe du coffre) ; connexion Microsoft interactive ; test de disponibilité |
| `Mail/MessageMetadata`, `MessageContent` | L'objet `email` plat envoyé à Jev (depuis le JSON Graph ou l'enveloppe/en-têtes IMAP) ; HTML→texte ; compaction tête/queue |
| `Storage/*` | Chemins, écritures JSON atomiques 0600, `AppConfig` (global), `RuleStore` (règles globales), `AccountStore` (`accounts.json`), `SecretStore`, `JobStore` (un par compte : session + règles figées) |
| `Graph/GraphAuth` | Credential par compte (tenant surchargeable), connexion device-code / navigateur, record sous `accounts/<id>/`, migration de l'ancien état mono-boîte |
| `Graph/GraphMailClient` | `IMailbox` sur Graph : listing par curseur `receivedDateTime`, lectures `$batch`, catégories maîtres (couleurs), PATCH des catégories, déplacement vers Archive, gestion du throttling |
| `Imap/ImapMailbox` | `IMailbox` sur MailKit : listing par UID, fetch enveloppe/en-têtes, labels Gmail ou mots-clés IMAP, archivage (All Mail / dossier Archive), une connexion sérialisée avec reconnexion |
| `Jev/JevPayloadBuilder` | Construit la question Choice (options `L0..Ln`), garde-fou de taille (29 000 octets), estimation de coût |
| `Jev/JevClient` | Envoie une requête Decisions ; validation stricte du contrat ; source du coût |
| `Triage/TriageEngine` | Cycle de vie de session, lots, deux vagues, concurrence adaptative, budget, disjoncteur, écritures groupées — indépendant du fournisseur |
| `Cli/*` | Commandes `account`, `--account` / `--all-accounts`, run/continue/status, rendu terminal, saisie masquée du mot de passe ; `ServiceCommand` (`service install/status/restart/uninstall`, LaunchAgent macOS, voir §10) |
| `Web/UiServer` | Tableau de bord local (ASP.NET Core minimal API sur 127.0.0.1) : sert la page embarquée et expose les fonctions serveur de jevMail en `POST /api/{function}` (arguments en tableau JSON) plus `addAccount` / `removeAccount` / `testAccount` ; un moteur et un état de connexion device-code par compte. `GET /health` → `{app, version, https}` identifie un jevOutlook en cours d'exécution (instance unique, `service status`) ; `AllowedOrigins(port, httpsPort)` est l'unique endroit qui construit les listes blanches Host/Origin |
| `Web/wwwroot/*` | `index.html` + `style.css` adaptés de jevMail (MIT) : un pont émule `google.script.run` sur `fetch` ; carte Mailboxes (puces, ajout/test/connexion/suppression, file « Run all mailboxes ») ; `taskpane.html` (Office.js) et icônes du complément Outlook |
| Complément Outlook | Manifeste XML classique servi sur `/manifest.xml` ; le volet est servi en HTTPS par le serveur local, que le LaunchAgent (§10) maintient en marche. Un tenant d'entreprise a refusé le chargement (« installation failed » générique alors que le manifeste passe le validateur Microsoft) : les tenants qui désactivent les compléments personnalisés le bloquent toujours. Un id de compte vide côté API désigne la première boîte Microsoft prête |

## 4. Correspondance par fournisseur

| Concept | Outlook (Graph) | Gmail (IMAP) | IMAP générique (ex. Dovecot) |
| --- | --- | --- | --- |
| Libellé | Catégorie maître (`POST /me/outlook/masterCategories` avec couleur) + `PATCH /me/messages/{id}` `{categories}` (remplacement complet → catégories existantes fusionnées, jamais perdues) | Label Gmail (dossier de l'espace personnel, créé s'il manque) posé via `X-GM-LABELS` (`AddLabels` / `RemoveLabels`) | Mot-clé IMAP (flag personnalisé) via `STORE +FLAGS` quand le dossier annonce `PERMANENTFLAGS \*` ; sinon l'étiquetage échoue fermé avec un message clair |
| Archivage | `POST /me/messages/{id}/move` `{destinationId:"archive"}` | Déplacement Inbox → `[Gmail]/All Mail` (= retrait du label Inbox ; le message reste dans All Mail) | Déplacement vers le dossier `Archive` (special-use ou créé) |
| Marqueur | aucun : « déjà traité » = porte l'un des libellés configurés | idem | idem |
| Portée `inbox` / `all` | `/me/mailFolders/inbox/messages` / `/me/messages` avec exclusion côté client de Junk, Deleted Items, Drafts | `INBOX` / `[Gmail]/All Mail` | `INBOX` seulement (`SupportsAllScope=false`, refusé au démarrage) |
| Non lus | `isRead eq false` | `NOT SEEN` | `NOT SEEN` |
| Id de message | Id Graph (change au déplacement → 404 au replay géré) | `scope:uidvalidity:uid` | idem |
| Clé de tri / curseur | `receivedDateTime` ISO-8601 (7 décimales, `Z`) | `uidvalidity:uid` complété à 12 chiffres (ordre lexicographique = numérique) | idem |
| Lecture métadonnées | `$batch` de `GET …?$select=…,internetMessageHeaders` (20 par batch) | un `FETCH` par fenêtre : `ENVELOPE`, `FLAGS`, `X-GM-LABELS`, en-têtes choisis, structure du corps, extrait texte de 2 Ko | idem sans labels |
| Lecture complète | `$select=id,body` avec `Prefer: outlook.body-content-type="text"` | `FETCH BODY.PEEK[]` → `TextBody` ou `HtmlBody`→texte | idem |
| Estimation | Compteurs de la boîte de réception ; `$count=true` pour toute la boîte | `Count` / `Unread` du dossier après `SELECT` | idem |
| Limites de débit | 429 `ApplicationThrottled` (MailboxConcurrency = 4 par boîte ; Graph exécute un `$batch` 4 par 4) → exactement **un `$batch` en vol**, fenêtre 5→20 ; lectures réessayées 3× en honorant Retry-After (≤ 30 s) ; écritures envoyées lot par lot et réessayées 5× (≤ 30 s) | connexion unique, sérialisée ; perte de connexion → pause transitoire | idem |
| Connexion | code device / navigateur, cache MSAL + record par compte | mot de passe d'application (validation en deux étapes) dans le Trousseau | mot de passe de la boîte dans le Trousseau |

## 5. Curseur de listing

Le moteur liste du plus récent au plus ancien à partir d'une **clé de tri** propre au
fournisseur et n'avance le curseur que sur les messages effectivement **examinés** : une page
peut contenir plus de candidats qu'un lot n'en consomme sans rien perdre. Les ids partageant
la clé du curseur sont gardés dans `CursorBoundaryIds` et exclus côté client (les ex æquo ne
sont ni répétés ni perdus) ; une page entière d'ids frontière bascule la requête suivante en
exclusif (`lt`). Le curseur étant une clé et non un jeton de page, archiver des messages hors
de la boîte de réception en cours de run ne décale jamais les pages.

- Graph : `$filter=receivedDateTime le {curseur}[ and isRead eq false]&$orderby=receivedDateTime desc&$top=N`
  (Graph exige que les propriétés de `$orderby` apparaissent en premier dans `$filter`). Curseur initial = début + 5 min.
- IMAP : `UID SEARCH 1:{uidCurseur}` (`uidCurseur-1` en exclusif) `[NOT SEEN]`, tri par UID
  décroissant, puis `FETCH` de la page. Curseur initial = `UIDNEXT`. Un changement d'`UIDVALIDITY`
  met la session en erreur avec un message clair (les UID n'ont plus de sens après une reconstruction).
- Les messages reçus après le démarrage de la session n'en font pas partie.

## 6. Algorithme de traitement (un lot)

1. **Remplir la file** (≤ 50) depuis le curseur ; terminé quand épuisé ou limite atteinte.
2. **Lecture des métadonnées** (fenêtre adaptative). Les éléments portant déjà un libellé
   configuré sans décision finale sont retirés (traités ailleurs). Introuvable → ignoré sans risque.
3. **Vague 1 — métadonnées** : une requête Choice indépendante par élément sans résultat.
   Les réponses réussies sont checkpointées avant le traitement des échecs.
4. **Décision** : confiance ≥ seuil métadonnées (et, en mode archivage pour un libellé
   éligible, ≥ seuil d'archivage) → finale ; sinon l'élément passe en vague 2.
5. **Lecture complète** + extraction du texte. Corps vide → repli sûr `review` (sans archivage)
   si une telle règle existe, sinon ignoré.
6. **Vague 2 — corps complet** : requêtes indépendantes ; réponses checkpointées.
7. **Préfixe prêt** : les éléments de tête avec décision finale sont écrits en deux passes
   groupées idempotentes — les libellés courants de chaque message sont **relus au même
   instant** (une écriture de remplacement sur un instantané vieux de plusieurs minutes
   écraserait les changements de l'utilisateur), puis `ApplyLabels` (courant ∪ {libellé}), puis
   `Archive` pour les décisions d'archivage. Un échec transitoire de la boîte met la session en
   pause avec toutes les décisions sauvegardées. Au replay, un élément avec décision finale
   mais introuvable (son id a changé parce que le déplacement a réussi avant le checkpoint) est
   compté comme traité, pas ignoré.
8. **Compteurs** (traités, métadonnées seules, complet, archivés, par libellé), tableau des
   résultats, file tronquée, session sauvegardée. Blocage budgétaire → statut `budget`.

### Envoi des vagues (`DispatchWaveAsync`)

- Réservation budgétaire avant envoi (`SpentUsd + réserve + estimation ≤ MaxSpendUsd`),
  remplacée ensuite par le coût réel (`reported` > `input_tokens` > `estimated`).
- Requêtes concurrentes par fenêtres de `JevConcurrency` (25 → 50, divisé par deux sur échec).
- Échecs de session réessayables (429/5xx/transport) : si toute la fenêtre a échoué avec le
  même code, **sonder une requête** avant de réessayer le reste ; sinon réessayer une fois les échecs.
- Échecs de contrat par message (JSON invalide, option inconnue, probabilités incohérentes) :
  un réessai, borné par le disjoncteur (3 consécutifs → pause).

### Politique d'échec (fermeture sûre)

| Situation | Effet |
| --- | --- |
| Fournisseur 401/402/403/4xx | Erreur / pause de session, aucune modification de la boîte |
| Fournisseur 429/5xx après réessai | Pause `provider-temporary`, reprenable. Un 429 n'est pas facturé (réservation libérée) ; timeouts/5xx gardent l'estimation prudente |
| Réponse modèle invalide ×1 | Réessai ; puis message ignoré, `ModelResponseSkips++` |
| 3 réponses invalides consécutives | Pause `jev-response-circuit-breaker` |
| Boîte 429/5xx / perte de connexion IMAP après réessais (lectures 3×, écritures 5×, Retry-After honoré) | Pause `mailbox-temporary` (lectures, relecture pré-écriture ou écritures groupées) ; `continue` reprend |
| Boîte 401/403 / authentification IMAP refusée | Erreur, demande de reconnexion (`account login`) ou de nouveau mot de passe (`account password`) |
| Message introuvable | Ignoré sans risque (lectures) ; considéré comme fait (écritures) ; compté comme traité quand l'élément porte déjà une décision finale (replay après déplacement) |
| 40 pages de listing sans nouveau candidat | Pause `scan-guard` (tout porte déjà un libellé configuré) ; `continue` poursuit le balayage |
| Ctrl+C pendant une vague | Réservations des requêtes sans réponse libérées ; réenvoyées à la reprise |
| Budget dépassé | Statut `budget` ; `continue --max-spend` pour relever |
| Ctrl+C | Pause `user-stop` ; `continue` reprend (et réarme le disjoncteur) |

## 7. Persistance

```
~/.jevoutlook/
  config.json            client id, tenant, fournisseur, surcharges endpoint/modèle, clé API optionnelle, flag device-code
  rules.json             règles sauvegardées (partagées par toutes les boîtes)
  accounts.json          [{id, email, kind: graph|imap, host, port, gmail, username, tenantId}] — aucun secret
  secrets.json           mots de passe IMAP, hors macOS seulement (macOS : Trousseau, service « jevoutlook »)
  accounts/<id>/
    auth-record.json     AuthenticationRecord MSAL (boîtes Microsoft ; les jetons vivent dans le cache de l'OS)
    job.json             session courante : options, curseur (clé de tri), compteurs, dépense, éléments en attente avec décisions
    job-rules.json       règles figées pour la session courante
  logs/ui.log            stdout/stderr du LaunchAgent
~/Library/LaunchAgents/com.vincentlauriat.jevoutlook.plist   LaunchAgent (macOS), aucun secret
```

`<id>` dérive de l'adresse (`alice@contoso.com` → `alice-contoso.com`). Tous les
fichiers sont écrits atomiquement (`.tmp` + move) en mode 0600 ; les répertoires sont en 0700.
L'ancienne disposition mono-boîte (`auth-record.json`, `job.json` à la racine) est migrée une
fois en entrée de compte au premier démarrage ; l'ancienne session est abandonnée car le type
du curseur a changé.

## 8. Notes de sécurité

- Microsoft : permissions déléguées uniquement (`Mail.ReadWrite`, `MailboxSettings.ReadWrite`,
  `User.Read`) ; l'application n'envoie jamais de courrier. IMAP : le mot de passe n'est transmis
  qu'à MailKit ; il n'est jamais journalisé.
- Les mots de passe sont stockés dans le Trousseau macOS via `/usr/bin/security` avec `ArgumentList`
  (jamais via un shell) et ne sont jamais écrits dans `accounts.json` ni `config.json`. Gmail exige
  un **mot de passe d'application** (validation en deux étapes), jamais le mot de passe du compte.
- Les corps d'erreur du fournisseur ne sont jamais réaffichés (ils peuvent contenir des données de requête).
- Le contenu des e-mails est transmis au modèle comme donnée non fiable ; les instructions lui
  interdisent de suivre des instructions trouvées dans l'e-mail.
- La clé API est lue d'abord dans l'environnement ; le stockage sur disque est optionnel.
- Le tableau de bord n'écoute que sur 127.0.0.1 et rejette le DNS rebinding et les appels cross-site :
  liste blanche `Host` (421), `Origin` / `Sec-Fetch-Site` de même origine et `application/json`
  obligatoires sur `/api/*` (403). Les mots de passe saisis dans la page ne vont qu'à ce serveur local.
  `GET /health` passe par le même contrôle de Host.
- Le plist du LaunchAgent est en clair : `BuildLaunchAgentPlist` n'écrit que `DOTNET_ROOT` et `JEVOUTLOOK_HOME`
  (liste blanche), jamais `OPENROUTER_API_KEY` / `JEV_API_KEY` ; l'agent lit la clé stockée avec `key set`.
  `launchctl` et `id -u` sont lancés avec `ArgumentList`, jamais via un shell.

## 9. État de vérification (2026-09-28)

- 2026-09-24 : `dotnet build` : 0 avertissement, 0 erreur. `dotnet test` : 66 tests passent (dont 18 tests des helpers IMAP ; MailKit 4.18.0).
- Revue de code indépendante (agent séparé) + vérification documentaire des faits Graph / Azure.Identity /
  TypeSafe le 2026-09-22 ; constats majeurs corrigés.
- En réel, sur une boîte Microsoft 365 professionnelle : connexion device-code, listing, lectures `$batch`, les deux étapes Jev,
  PATCH de catégories (20 messages) et nettoyage de catégorie ; après la refonte multi-comptes, l'état
  existant a migré en place et un run preview de 10 messages depuis le tableau de bord a abouti (0,00067 $).
- 2026-09-24, run réel de 100 messages de la boîte de réception de cette boîte Microsoft 365 (catégories seules) : la première tentative s'est
  mise en pause deux fois sur des 429 `ApplicationThrottled` de Graph (deux `$batch` en vol = 8 requêtes simultanées >
  MailboxConcurrency 4 ; la relecture pré-écriture n'avait aucun réessai). Après correction (un lot en vol, relecture
  adaptative, réessais honorant Retry-After) le run s'est terminé : 99 catégorisés, 1 ignoré (pas de texte lisible),
  120 requêtes Jev, 0,009 $, zéro 429.
- Pas encore exercé en réel : le déplacement vers Archive ; Gmail et IMAP générique (aucun identifiant
  disponible : mots de passe d'application Gmail, mot de passe IMAP, consentement administrateur du second tenant Microsoft).
- 2026-09-28 : `dotnet build -c Release` 0 avertissement, 0 erreur ; `dotnet test` 79 tests passent (+13 : contenu du plist
  du LaunchAgent, exclusion des secrets, échappement XML, `plutil -lint`, liste blanche des origines, reconnaissance du corps `/health`).

## 10. Maintenir le serveur en marche (LaunchAgent macOS)

Le volet du complément Outlook est servi par `jevoutlook ui` lui-même, et un complément web Office ne peut pas lancer de processus local.
Décision du 2026-09-28 (option 1) : un LaunchAgent utilisateur maintient le serveur en marche dès l'ouverture de session.

| Élément | Comportement |
| --- | --- |
| `service install [--port] [--https-port] [--exe]` | Écrit `~/Library/LaunchAgents/com.vincentlauriat.jevoutlook.plist` (`ProgramArguments` = exe `ui --no-open --port --https-port`, `RunAtLoad`, `KeepAlive`, `ThrottleInterval` 30, stdout/stderr → `~/.jevoutlook/logs/ui.log`), puis `launchctl bootout` (si chargé) + `bootstrap gui/<uid>`. L'exécutable par défaut est celui en cours ; avertit pour `bin/Debug`/`bin/Release` et recommande une copie `dotnet publish` dans `~/.jevoutlook/bin` |
| `service status` | Plist présent, agent chargé (`launchctl print`), sonde `/health`, chemin du journal ; code de sortie 1 si le serveur ne répond pas |
| `service restart` | `launchctl kickstart -k gui/<uid>/<label>` (après une nouvelle publication) |
| `service uninstall` | `launchctl bootout`, puis supprime le plist |
| Instance unique | `ui` sonde d'abord `http://127.0.0.1:{port}/health` (2 s) : réponse jevOutlook → « already running », ouverture du navigateur, sortie 0 ; toute autre réponse → erreur claire au lieu de l'exception de bind de Kestrel |

Comportement connu : tant qu'un `ui` lancé à la main occupe le port, la copie de l'agent sort en 0 (« already running ») et launchd,
avec `KeepAlive` à true, la relance toutes les 30 s (une ligne de journal à chaque fois). `KeepAlive = {SuccessfulExit: false}`
l'éviterait, au prix de ne plus relancer après une sortie propre.

L'option 2 (démarrage à la demande depuis un volet hébergé statiquement, via un gestionnaire d'URL `jevoutlook://start`) est
conçue, pas construite : voir [docs/design/addin-on-demand-start.md](docs/design/addin-on-demand-start.md).
