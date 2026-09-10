# Business proposal

A proposal asks a client to choose you, for this project, at this price. Everything in it
serves that decision. There is no ISO standard for proposals; the structure below is common
practice in professional services, and the writing rules in `writing.md` apply throughout.

Template: `templates/proposal.md`.

## Structure

| Section | What it answers |
| :-- | :-- |
| Executive summary | What the client wants, what you propose, why you. Half a page; the only part some readers read. |
| About us | Why you are credible for this project: similar work, not a company history. |
| Project overview | Your understanding of the client's situation, objectives and constraints. |
| Scope of work | Exactly what is included, as a table of modules and features, and what is not. |
| Solution | How it will be built: technology, architecture diagram. |
| Deliverables | What the client receives: software, documents, training. |
| Team | Roles, headcount, responsibilities. |
| Timeline | Phases and milestones, as a chart with real dates. |
| Assumptions and dependencies | What must be true for the price and dates to hold, and who provides what. |
| Price and payment | Price per item, total, tax treatment, payment milestones. |
| Terms and conditions | Validity, change requests, warranty, intellectual property. |

## Rules

- **Write the executive summary last**, for this client. A summary that could be sent to
  anyone persuades no one.
- **Scope is a table**, one feature per row, with an In scope column. Add the sentence:
  anything not listed is out of scope and is a change request.
- **Every price states its tax treatment** ("excluding VAT") under the table, and the
  currency is in the column header.
- **Dependencies are named with an owner**. Never promise a third party's or the client's
  work as your deliverable; put it under assumptions and dependencies.
- **The timeline starts from an event you control**, such as requirement sign-off, and says so.
- **Customer-facing language only**: no internal notes, no criticism of the client or of
  other vendors, no pressure on the client's deadlines.
- **Validity**: state how long the offer stands (60 days is common).
- **Confidentiality notice** on the cover.

## Before sending

- Every `{{placeholder}}` filled (`--strict` fails the build otherwise).
- Scope, team size, timeline and price agree with the internal estimate.
- The client's legal name is spelled exactly as registered.
- Render the PDF and read every page.
