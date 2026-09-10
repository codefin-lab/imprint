---
# A technical specification in the arc42 structure (ISO/IEC/IEEE 42010 concepts, C4 diagrams).
# Guidance: skills/docgen/references/tech-spec.md. Build:
#   python3 build_docx.py templates/tech-spec.md --pdf
document_kind: Technical Specification
PROJECT_NAME: "Example Project: Onboarding Platform"
CLIENT_LEGAL_NAME: "Example Client Co., Ltd."
VERSION: "0.1"
DD MMM YYYY: "1 Oct 2026"
Draft / For Review / Final: "Draft"
---

# DOCUMENT CONTROL

| Version | Date | Author | Description of change |
| :-- | :-- | :-- | :-- |
| 0.1 | 2026-10-01 | Tech lead | First draft |

# TABLE OF CONTENTS

<!-- toc -->

# 1. INTRODUCTION AND GOALS

## Requirements overview

The platform lets customers open an account from the mobile app or at a branch, and gives staff one record per customer. The business requirements are in the BRD, version 1.0.

## Quality goals

| Priority | Quality | Goal |
| --: | :-- | :-- |
| 1 | Security | Personal data is protected in transit, at rest and in logs |
| 2 | Performance efficiency | 95% of submissions complete within 2 s at peak |
| 3 | Maintainability | A new product type is added without changing the core services |

## Stakeholders

| Role | Expectation of this document |
| :-- | :-- |
| Development team | What to build and why |
| Operations | What to run, where, and how to monitor it |
| Security reviewer | How data and access are protected |

# 2. CONSTRAINTS

| Constraint | Background |
| :-- | :-- |
| Customer data is hosted in-country | Regulatory requirement |
| Integration with the existing core system through its REST API only | The core system is owned by another team |

# 3. CONTEXT AND SCOPE

The platform, its users and the systems around it. Draw it as a C4 context diagram with the diagram-design skill.

| Neighbour | Interface | Direction |
| :-- | :-- | :-- |
| Mobile app | HTTPS, JSON | Inbound |
| National ID service | HTTPS, JSON | Outbound |
| Core system | REST API | Outbound |

# 4. SOLUTION STRATEGY

- Services split by business capability, behind one API gateway
- Asynchronous events between services, so a slow neighbour does not block onboarding
- Managed database with automated backups and point-in-time recovery

# 5. BUILDING BLOCKS

![Figure 1: Containers of the onboarding platform](example-architecture.png)

| Container | Responsibility | Technology |
| :-- | :-- | :-- |
| API gateway | Authentication, routing, rate limiting | Your gateway |
| Customer service | Customer record and identity verification | Your stack |
| Order service | Applications and their status | Your stack |
| Database | Primary store | Your database |

# 6. RUNTIME VIEW

## Submit an application

1. The app sends the application to the API gateway with the customer's token.
2. The gateway checks the token and forwards the request to the order service.
3. The order service stores the application and publishes an ApplicationSubmitted event.
4. The customer service verifies identity and publishes the result.

Add a sequence diagram for each scenario that matters.

# 7. DEPLOYMENT VIEW

| Environment | Node | vCPU | RAM (GB) | Storage (GB) | Count |
| :-- | :-- | --: | --: | --: | --: |
| Production | Application | 4 | 16 | 100 | 3 |
| Production | Database | 8 | 32 | 500 | 2 |
| UAT | Application | 2 | 8 | 50 | 2 |

# 8. CROSS-CUTTING CONCEPTS

| Concept | Approach |
| :-- | :-- |
| Authentication | OAuth 2.0 access tokens issued by the identity provider |
| Secrets | Held in a secrets manager, never in code or images |
| Logging | Structured logs without personal data, retained 90 days |
| Monitoring | Health checks, latency and error-rate alerts per service |

# 9. ARCHITECTURE DECISIONS

| ID | Decision | Status | Date |
| :-- | :-- | :-- | :-- |
| ADR-001 | Use an event bus between services | Accepted | 2026-10-01 |
| ADR-002 | Use a managed database service | Accepted | 2026-10-01 |

# 10. QUALITY REQUIREMENTS

| ID | Quality | Scenario | Measure |
| :-- | :-- | :-- | :-- |
| QS-1 | Performance efficiency | 200 customers submit at the same time | 95% complete within 2 s |
| QS-2 | Reliability | One application node fails | No failed submissions; recovery within 1 min |

# 11. RISKS AND TECHNICAL DEBT

| ID | Risk | Mitigation |
| :-- | :-- | :-- |
| R-01 | The ID service has no load-test environment | Agree a test window and a rate limit with the provider |

# 12. GLOSSARY

| Term | Meaning |
| :-- | :-- |
| C4 | A model for software architecture diagrams at four levels: context, container, component, code |
