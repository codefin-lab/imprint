---
# A business requirements document (ISO/IEC/IEEE 29148, IIBA BABOK).
# Guidance: skills/docgen/references/brd.md. Build:
#   python3 build_docx.py templates/brd.md --pdf
document_kind: Business Requirement Document
PROJECT_NAME: "Example Project: Customer Onboarding"
CLIENT_LEGAL_NAME: "Example Client Co., Ltd."
VERSION: "0.1"
DD MMM YYYY: "1 Oct 2026"
Draft / For Review / Final: "Draft"
---

# DOCUMENT CONTROL

| Version | Date | Author | Description of change |
| :-- | :-- | :-- | :-- |
| 0.1 | 2026-10-01 | Business analyst | First draft |

| Role | Name | Position | Date |
| :-- | :-- | :-- | :-- |
| Preparer |  | Business analyst |  |
| Reviewer |  | Product owner |  |
| Approver |  | Sponsor |  |

# TABLE OF CONTENTS

<!-- toc -->

# 1. INTRODUCTION

## Purpose

This document states the business requirements for {{PROJECT_NAME}}. It is the basis for design, build and acceptance testing, and changes only through change control once approved.

## Conventions

"Shall" marks a binding requirement, "should" a recommendation, "may" an option. Requirement IDs are stable and never reused.

# 2. BUSINESS CONTEXT

## Problem

A new customer waits five working days to open an account, because details are captured three times across three systems.

## Objectives

| ID | Objective | Measure of success |
| :-- | :-- | :-- |
| OBJ-1 | Shorten onboarding | Median time from application to active account under 1 working day |
| OBJ-2 | Capture customer data once | No field re-keyed after the application form |

# 3. STAKEHOLDERS AND USERS

| Stakeholder | Interest | Needs |
| :-- | :-- | :-- |
| Customer | Opens an account | Apply once, from a phone, and know the status |
| Branch staff | Serve walk-in customers | One screen with the customer's complete record |
| Compliance | Meets KYC rules | Every identity check recorded and auditable |

# 4. SCOPE

## In scope

- Account opening in the mobile app and at the branch
- Identity verification and KYC screening

## Out of scope

- Loan and card applications
- Migration of historical customer records

# 5. BUSINESS PROCESS

Describe the current process, then the future process. Add a diagram of each, drawn with the diagram-design skill.

# 6. FUNCTIONAL REQUIREMENTS

| ID | Requirement | Priority | Source | Acceptance criteria |
| :-- | :-- | :-: | :-- | :-- |
| FR-001 | The system shall let a customer submit an account application from the mobile app. | Must | OBJ-1 | Given a new customer, when they submit a complete form, then an application is created and its reference shown. |
| FR-002 | The system shall verify the customer's identity against the national ID service. | Must | OBJ-1 | Given a submitted application, when verification succeeds, then the application moves to KYC screening. |
| FR-003 | The system shall show branch staff the full application in one screen. | Should | OBJ-2 | Given an application, when staff open it, then every captured field is visible without switching systems. |

# 7. NON-FUNCTIONAL REQUIREMENTS

| ID | Quality | Requirement | Priority |
| :-- | :-- | :-- | :-: |
| NFR-001 | Performance efficiency | The application form shall submit within 2 s for 95% of requests at 200 concurrent users. | Must |
| NFR-002 | Security | Personal data shall be encrypted in transit and at rest. | Must |
| NFR-003 | Reliability | The onboarding service shall be available 99.5% of each calendar month. | Should |

# 8. TRANSITION REQUIREMENTS

| ID | Requirement | Priority |
| :-- | :-- | :-: |
| TR-001 | Branch staff shall be trained before go-live. | Must |

# 9. ASSUMPTIONS, CONSTRAINTS AND DEPENDENCIES

| Type | Statement | Impact if it does not hold |
| :-- | :-- | :-- |
| Assumption | The national ID service provides a test environment. | Identity verification cannot be tested before go-live. |
| Constraint | Customer data stays in-country. | Hosting options are limited to local regions. |

# 10. RISKS

| ID | Risk | Likelihood | Impact | Owner | Mitigation |
| :-- | :-- | :-- | :-- | :-- | :-- |
| R-01 | The ID service changes its interface during the project | Low | High | Tech lead | Agree a version freeze with the provider |

# 11. GLOSSARY

| Term | Meaning |
| :-- | :-- |
| KYC | Know your customer: the identity and risk checks required before opening an account |
