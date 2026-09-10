---
# Technical report or spec: architecture, sizing, assessment.
#   python3 build_docx.py examples/report.md --pdf
cover: false
header_left: Technical Report
page_break_before_h1: false   # a report flows; use \pagebreak where you want one
PROJECT_NAME: "{{Project}}"
VERSION: "1.0"
DD MMM YYYY: "{{DD MMM YYYY}}"
REPORT_TITLE: "{{Service Architecture & Infrastructure Sizing}}"
CLIENT_SHORT_NAME: "{{Client}}"
PROJECT_START: 2026-10          # the chart needs real dates, not placeholders
KICKOFF_DATE: 2026-10-06
---

# {{REPORT_TITLE}}

**{{PROJECT_NAME}}: {{COMPANY_NAME}} for {{CLIENT_SHORT_NAME}}**
Version {{VERSION}}  |  {{DD MMM YYYY}}

# 1. Executive summary

{{What this document answers, in the terms the reader asked. If it responds to
specific action items or questions, name them here.}}

| # | Question / action item | Answered in |
| --: | :-- | :-- |
| 1 | {{Action item}} | Section {{n}} |

# 2. Architecture

{{Prose that carries the shape of the system. Insert the diagram as an image
only once it exists; an empty heading reads worse than no heading.}}

## 2.1 Service groups

| Group | Service | Responsibility |
| :-- | :-- | :-- |
| **{{Group}}** |  |  |
| | {{service}} | {{what it does}} |

## 2.2 Dependencies

| Service | Depends on | Kind | Notes |
| :-- | :-- | :-- | :-- |
| {{service}} | {{service}} | {{sync / async / batch}} | {{note}} |

# 3. Sizing

{{State the workload the numbers rest on before the numbers. A sizing without
its assumptions cannot be argued with, and will be wrong later.}}

| Environment | Node | vCPU | RAM | Storage | Count |
| :-- | :-- | --: | --: | --: | --: |
| {{PROD}} | {{node}} | {{4}} | {{16 GB}} | {{100 GB}} | {{3}} |

## 3.1 Assumptions the sizing rests on

1. {{Users, peak concurrency, transactions per day.}}
2. {{Retention and growth.}}
3. {{What is explicitly out of the estimate.}}

<!-- landscape -->

# 4. Timeline

```markwhen
title: {{PROJECT_NAME}}, indicative plan

group Design
{{PROJECT_START}}/2 months: Requirements and design #design
endGroup

{{KICKOFF_DATE}}: Kick-off
```

<!-- portrait -->

# 5. What we need from {{CLIENT_SHORT_NAME}}

1. {{The decision or input that unblocks the next step, and by when.}}

# 6. Open items

- {{Anything still unresolved. Say it here rather than leaving it implied.}}

---

*Confidential. Copyright {{COMPANY_NAME}}. Prepared for
{{CLIENT_SHORT_NAME}}. All rights reserved.*
