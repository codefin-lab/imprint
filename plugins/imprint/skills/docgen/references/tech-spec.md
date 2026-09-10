# Technical specification

A technical specification says how a system is built: its parts, how they talk, where they
run, and the decisions behind them, so that a team can build it and a reviewer can judge it.

The structure follows **arc42**, the widely used open template for architecture
documentation, and the concepts of ISO/IEC/IEEE 42010:2022 (architecture description:
stakeholders, concerns, views). Diagrams follow the **C4 model** levels: context, container,
component. Template: `templates/tech-spec.md`.

## Structure (arc42)

| # | Section | Contents |
| --: | :-- | :-- |
| 1 | Introduction and goals | requirements overview, top 3 to 5 quality goals, stakeholders |
| 2 | Constraints | technical, organisational and regulatory constraints the design must accept |
| 3 | Context and scope | the system as a box, its users and neighbouring systems (C4 context diagram) |
| 4 | Solution strategy | the few decisions that shape everything else |
| 5 | Building blocks | the containers and components, what each is responsible for (C4 container diagram) |
| 6 | Runtime view | how the parts work together in the important scenarios (sequence diagrams) |
| 7 | Deployment view | environments, infrastructure, sizing |
| 8 | Cross-cutting concepts | security, identity, logging, monitoring, error handling, data protection |
| 9 | Architecture decisions | the decision records |
| 10 | Quality requirements | quality scenarios that make the quality goals testable |
| 11 | Risks and technical debt | known risks and the plan for each |
| 12 | Glossary | terms and abbreviations |

Leave a section out only if it truly does not apply, and say so in one line rather than
deleting the heading.

## Rules

- **Diagrams at C4 levels**, one level per diagram, every box and arrow labelled with what
  it is and how it talks (protocol, sync or async).
- **Sizing is a table** with figures right-aligned and units in the header:
  `| Environment | Node | vCPU | RAM (GB) | Storage (GB) | Count |`.
- **Quality goals are measurable** and each has a scenario in section 10: source, stimulus,
  environment, response, measure. "Under a load of 500 orders per minute, 95% of orders are
  confirmed within 2 s."
- **Decisions are recorded as ADRs** (architecture decision records): title, status, context,
  decision, consequences. One decision per record; a superseded record stays, marked Superseded.
- **Security is explicit**: authentication, authorisation, secrets, encryption in transit and
  at rest, audit logging, personal-data handling.
- **Every external dependency is named** with its owner and what happens when it is down.

## Decision record table

| ID | Decision | Status | Date |
| :-- | :-- | :-- | :-- |
| ADR-001 | Use PostgreSQL as the primary store | Accepted | 2026-10-01 |
