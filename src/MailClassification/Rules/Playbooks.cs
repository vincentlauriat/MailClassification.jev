namespace MailClassification.Rules;

/// <summary>
/// Ready-to-adapt category sets. The first playbook is the default rule set.
/// Descriptions are the classification criteria sent to Jev verbatim.
/// </summary>
public static class Playbooks
{
    public static readonly IReadOnlyList<Playbook> All =
    [
        new Playbook(
            "universal",
            "Universal inbox",
            "A balanced starting point for personal and professional inboxes. It separates work, relationships, transactions, reading, automation, outreach, and genuinely unsafe mail.",
            [
                new LabelRule("action-required", "action-required", "A legitimate message that requires a reply, decision, approval, task, or time-sensitive intervention. Includes unresolved access, security, payment, delivery, legal, or account problems. Excludes optional reading and routine confirmations.", false),
                new LabelRule("important-update", "important-update", "A meaningful update from a trusted person, project, customer, employer, school, service, or account that should be retained, but does not currently require action. Excludes personal conversation, transactions, and mass editorial content.", false),
                new LabelRule("personal", "personal", "A person-to-person conversation with family, friends, or a known community contact whose main purpose is personal communication. Excludes business automation, transactional notices, and unsolicited commercial outreach.", false),
                new LabelRule("money-and-orders", "money-and-orders", "Receipts, invoices, payment or subscription confirmations, order and delivery updates, bank notices, and official financial documents when no unresolved problem requires action. Problems and disputes belong in action-required.", false),
                new LabelRule("opportunities", "opportunities", "A specific and credible new job, business, partnership, speaking, media, funding, or collaboration opportunity with relevant context. Excludes generic mass pitches, vague networking requests, and tasks from existing relationships.", false),
                new LabelRule("newsletters", "newsletters", "Opt-in recurring editorial, product, industry, creator, or community content intended for later reading. Excludes direct correspondence, account notices, and messages that require a reply or decision.", false),
                new LabelRule("routine-notifications", "routine-notifications", "Low-risk automated social, application, monitoring, digest, or system notifications that require no reply or action. Excludes security, access, payment, delivery, legal, and service-failure alerts.", true),
                new LabelRule("cold-outreach", "cold-outreach", "Unsolicited commercial, recruiting, PR, SEO, agency, sponsorship, or link-exchange outreach with no relevant active relationship or conversation. Excludes specific credible opportunities that clearly match the recipient.", true),
                new LabelRule("junk", "junk", "Phishing, scams, deceptive promotions, irrelevant mass spam, malicious requests, or incoherent disposable mail. Excludes merely low-priority but legitimate notifications and newsletters.", true),
                new LabelRule("review", "review", "A legitimate or ambiguous message that does not clearly fit any other configured label. Use this safe fallback instead of forcing an unsupported category.", false),
            ]),

        new Playbook(
            "founders-operators",
            "Founders & operators",
            "Prioritizes decisions, customer risk, investors, team blockers, delegation, and unwanted vendor outreach.",
            [
                new LabelRule("decision-required", "decision-required", "Requires the founder or operator to make a concrete decision, approval, trade-off, or commitment. Excludes FYI updates and tasks that can proceed without their judgment.", false),
                new LabelRule("customer-risk", "customer-risk", "Signals churn, a severe complaint, failed delivery, contractual concern, escalation, or revenue risk from a current customer. Excludes ordinary product questions and positive feedback.", false),
                new LabelRule("investor-and-board", "investor-and-board", "Communication from current investors, board members, or active fundraising counterparties. Excludes generic fundraising services and mass investor lists.", false),
                new LabelRule("team-blocker", "team-blocker", "A team member cannot continue meaningful work without input, access, approval, or conflict resolution from the recipient. Excludes routine status reports.", false),
                new LabelRule("delegatable", "delegatable", "A legitimate operational request that needs action but can reasonably be assigned to another owner without executive judgment.", false),
                new LabelRule("vendor-pitch", "vendor-pitch", "Unsolicited agency, software, consulting, SEO, recruiting, PR, or outsourcing outreach with no active buying process or existing relationship.", true),
                new LabelRule("review", "review", "A legitimate or ambiguous message that does not clearly match another founder and operator category.", false),
            ]),

        new Playbook(
            "sales-business-development",
            "Sales & business development",
            "Surfaces buying intent, deal work, partnerships, customer follow-up, sales operations, and irrelevant outreach.",
            [
                new LabelRule("hot-lead", "hot-lead", "A prospect shows specific buying intent by requesting pricing, a demo, procurement information, availability, scope, or a next step. Excludes generic interest and vendor pitches.", false),
                new LabelRule("deal-action", "deal-action", "An active opportunity needs a reply, proposal revision, security response, legal review, approval, or scheduled follow-up.", false),
                new LabelRule("partner-opportunity", "partner-opportunity", "A credible channel, integration, co-marketing, referral, reseller, or strategic partnership proposal with concrete mutual relevance.", false),
                new LabelRule("customer-success", "customer-success", "A current customer needs adoption help, renewal attention, issue resolution, or relationship follow-up. Excludes new-prospect conversations.", false),
                new LabelRule("sales-automation", "sales-automation", "Legitimate CRM, scheduling, call-recording, sequence, pipeline, or revenue-operations notifications related to the sales workflow.", false),
                new LabelRule("irrelevant-outreach", "irrelevant-outreach", "Unsolicited sales, recruiting, PR, or partnership outreach that is generic, mismatched, or unrelated to an active opportunity.", true),
                new LabelRule("review", "review", "A legitimate or ambiguous message that does not clearly match another sales and business development category.", false),
            ]),

        new Playbook(
            "investors",
            "Investors & funds",
            "Separates warm deal flow, active diligence, portfolio work, LP communication, ecosystem updates, and mass pitches.",
            [
                new LabelRule("founder-intro", "founder-intro", "A warm introduction or direct founder message with credible context, company details, and potential investment relevance. Excludes mass mailings.", false),
                new LabelRule("active-deal", "active-deal", "Communication about an opportunity already under evaluation, including diligence, documents, meetings, terms, references, or an investment decision.", false),
                new LabelRule("portfolio-action", "portfolio-action", "A portfolio founder or team needs a decision, introduction, assistance, escalation, or board-level response.", false),
                new LabelRule("lp-and-fund", "lp-and-fund", "Communication from limited partners, fund administrators, counsel, auditors, banks, or service providers about fund operations and reporting.", false),
                new LabelRule("ecosystem-update", "ecosystem-update", "Relevant market, sector, accelerator, demo-day, community, or founder update worth reviewing without immediate action.", false),
                new LabelRule("mass-fundraising-pitch", "mass-fundraising-pitch", "Generic fundraising blast, paid placement, brokered list, or startup promotion with little evidence of personal relevance.", true),
                new LabelRule("review", "review", "A legitimate or ambiguous message that does not clearly match another investor and fund category.", false),
            ]),

        new Playbook(
            "recruiting-people",
            "Recruiting & people operations",
            "Organizes candidate actions, interviews, offers, sensitive employee matters, recruiting systems, and vendor pitches.",
            [
                new LabelRule("candidate-action", "candidate-action", "A candidate needs a reply, review, decision, feedback, document, or next step from the recruiting team.", false),
                new LabelRule("interview-scheduling", "interview-scheduling", "Messages about interview availability, calendar coordination, rescheduling, interviewer assignments, or logistics.", false),
                new LabelRule("offer-and-closing", "offer-and-closing", "Communication about compensation, references, offer approval, negotiation, acceptance, start date, or preboarding.", false),
                new LabelRule("employee-sensitive", "employee-sensitive", "Confidential or high-impact employee relations, performance, leave, grievance, legal, payroll, or workplace safety communication.", false),
                new LabelRule("ats-automation", "ats-automation", "Legitimate applicant-tracking, assessment, background-check, scheduling, or onboarding system notifications.", false),
                new LabelRule("recruiting-vendor-pitch", "recruiting-vendor-pitch", "Unsolicited staffing agency, sourcing tool, employer-branding, benefits, or HR software outreach with no active evaluation.", true),
                new LabelRule("review", "review", "A legitimate or ambiguous message that does not clearly match another recruiting and people operations category.", false),
            ]),

        new Playbook(
            "freelancers-creators",
            "Freelancers, consultants & creators",
            "Highlights qualified work, active clients, contracts, community, platform notices, and low-quality collaboration spam.",
            [
                new LabelRule("qualified-opportunity", "qualified-opportunity", "A credible project, sponsorship, speaking, media, consulting, or collaboration request with relevant scope, timing, budget, or context.", false),
                new LabelRule("client-action", "client-action", "An active client needs a deliverable, decision, reply, revision, meeting, approval, or issue resolution.", false),
                new LabelRule("payment-and-contract", "payment-and-contract", "Invoices, payment status, tax documents, statements of work, contracts, signatures, licensing, or usage-rights communication.", false),
                new LabelRule("audience-and-community", "audience-and-community", "Meaningful audience replies, member questions, community moderation, event communication, or direct reader feedback.", false),
                new LabelRule("platform-update", "platform-update", "Legitimate notices from publishing, commerce, social, analytics, advertising, or creator platforms about account activity and performance.", false),
                new LabelRule("low-quality-collab", "low-quality-collab", "Generic guest-post, backlink, unpaid promotion, vague collaboration, exposure-only, or mass sponsorship outreach without credible fit.", true),
                new LabelRule("review", "review", "A legitimate or ambiguous message that does not clearly match another freelancer, consultant, or creator category.", false),
            ]),

        new Playbook(
            "support-commerce",
            "Support & e-commerce",
            "Routes escalations, billing, fulfillment, product help, customer insight, and system-generated support mail.",
            [
                new LabelRule("urgent-escalation", "urgent-escalation", "A customer reports safety, security, legal, repeated service failure, public escalation, severe business impact, or an imminent deadline.", false),
                new LabelRule("refund-or-billing", "refund-or-billing", "A request or dispute involving charges, refunds, invoices, subscriptions, payment failures, taxes, or billing details.", false),
                new LabelRule("delivery-or-order", "delivery-or-order", "Questions or problems involving an order, shipment, address, fulfillment, stock, cancellation, return, or delivery status.", false),
                new LabelRule("product-help", "product-help", "The customer needs instructions, troubleshooting, compatibility guidance, account help, or an explanation of product behavior.", false),
                new LabelRule("feedback-and-feature", "feedback-and-feature", "Product feedback, feature requests, usability observations, reviews, or research participation without an unresolved support issue.", false),
                new LabelRule("automated-system-mail", "automated-system-mail", "Legitimate ticket acknowledgements, monitoring notices, routing messages, and support-platform automation that do not contain a new customer request.", false),
                new LabelRule("review", "review", "A legitimate or ambiguous message that does not clearly match another support and e-commerce category.", false),
            ]),
    ];

    public static IReadOnlyList<LabelRule> DefaultRules => All[0].Rules;

    public static Playbook? Find(string id) =>
        All.FirstOrDefault(p => string.Equals(p.Id, id, StringComparison.OrdinalIgnoreCase));
}
