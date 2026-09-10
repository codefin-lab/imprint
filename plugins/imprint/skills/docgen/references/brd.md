# Business requirements document (BRD)

A BRD states what the business needs and why, in terms the business can approve and a team
can build and test against. It does not design the solution.

It follows ISO/IEC/IEEE 29148:2018 (requirements engineering: the business requirements
specification) and the requirement classes of the IIBA BABOK Guide v3. Template:
`templates/brd.md`.

## Classes of requirement (BABOK)

| Class | States | Example |
| :-- | :-- | :-- |
| Business | a goal or objective of the organisation | Reduce onboarding time from 5 days to 1 |
| Stakeholder | what a group of stakeholders needs to reach it | Branch staff need one screen for a customer's details |
| Solution, functional | a behaviour the solution shall have | The system shall verify identity against the national ID service |
| Solution, non-functional | a quality the solution shall have | Search results shall return within 2 s for 95% of requests |
| Transition | what is needed only to move from now to then | Existing customer records shall be migrated before go-live |

## Structure

1. Introduction: purpose, scope of this document, definitions, references
2. Business context: the problem, objectives with measurable KPIs, benefits
3. Stakeholders and users: who is affected and what each needs
4. Scope: in scope, out of scope
5. Business process: the current process and the future process, with a diagram of each
6. Business and stakeholder requirements
7. Functional requirements
8. Non-functional requirements, grouped by the ISO/IEC 25010 quality characteristics that
   matter here: performance efficiency, reliability, security, interaction capability,
   compatibility, maintainability, flexibility, safety, functional suitability
9. Transition requirements
10. Assumptions, constraints and dependencies
11. Risks: ID, risk, likelihood, impact, owner, mitigation
12. Glossary

## A good requirement (ISO/IEC/IEEE 29148)

Each one is **necessary, appropriate, unambiguous, complete, singular, feasible, verifiable,
correct and conforming**. In practice:

- One requirement per row, with a unique, stable ID that is never reused: `BR-001`, `FR-012`.
- "The <actor> shall <verb> <object> <condition>." Measurable where a number applies.
- No "and/or", "etc.", "user-friendly", "fast", "as appropriate": each hides a second
  requirement or an untestable one.
- A priority on every row. MoSCoW (Must, Should, Could, Won't this time) is the common scale.
- A source on every row: who asked for it, or which objective it serves.
- Acceptance criteria that a tester can run. For user stories, Given / When / Then.

Requirement table:

| ID | Requirement | Priority | Source | Acceptance criteria |
| :-- | :-- | :-: | :-- | :-- |
| FR-001 | The system shall ... | Must | OBJ-1 | Given ..., when ..., then ... |

## Traceability

Every functional requirement traces up to a business objective and down to a test. Keep the
objective ID in the Source column; a requirement that traces to nothing is a candidate for
removal.

## Change

Once approved, the BRD is changed only through change control: a new version, a line in the
version history, and approval again.
