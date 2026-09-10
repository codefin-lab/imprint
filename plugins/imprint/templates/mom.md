---
# Minutes of meeting.  Fill these in, then:
#   python3 build_docx.py examples/mom.md --pdf
cover: false
header_left: Minutes of Meeting
page_break_before_h1: false   # a minute flows; it is not one section per page
PROJECT_NAME: "{{Project}}"
VERSION: "1.0"
DD MMM YYYY: "{{DD MMM YYYY}}"
MEETING_TITLE: "{{Company and client: subject}}"
MEETING_DATE: "{{DD Month YYYY}}"
MEETING_TIME: "{{HH:MM}}"
MEETING_DURATION: "{{1 h 30 m}}"
---

# MINUTES OF MEETING

**{{MEETING_TITLE}}**
{{MEETING_DATE}}  |  {{MEETING_TIME}}  |  {{MEETING_DURATION}}

# 1. Meeting information

| Item | Detail |
| :-- | :-- |
| Topic | {{MEETING_TITLE}} |
| Date | {{MEETING_DATE}} |
| Start | {{MEETING_TIME}} |
| Duration | {{MEETING_DURATION}} |
| Channel | {{Teams, on-site, or phone}} |
| {{COMPANY_NAME}} attendees | {{names, roles}} |
| Client attendees | {{names, roles}} |

**Purpose**

{{One or two sentences: what this meeting was called to settle.}}

# 2. Executive summary

- {{The decision or position that matters most, stated plainly.}}
- {{Who owns what, at the level a reader who missed the meeting needs.}}
- {{What is urgent and why.}}
- {{What the next meeting has to settle.}}

# 3. Discussion

## 3.1 {{Topic}}

- {{Point raised, and by whom when it matters.}}
- {{Position taken, with the reason.}}
- {{Anything left open: say so here rather than implying agreement.}}

## 3.2 {{Topic}}

- {{Point}}

# 4. Decisions and agreements

| # | Decision | Agreed by |
| --: | :-- | :-- |
| 1 | {{What was actually agreed, not what was discussed}} | {{Both parties, us, or the client}} |
| 2 |  |  |

# 5. Action items

| # | Action item | Owner | Due / status |
| --: | :-- | :-- | :-- |
| 1 | {{Action}} | {{Name}} | {{Date}} |
| 2 |  |  |  |

# 6. Open items to follow up

- {{Question still unanswered, and who has to answer it.}}

---

*These minutes record the understanding of the meeting as taken by {{COMPANY_NAME}}.
Please raise any correction within {{5}} business days; otherwise they are
taken as agreed.*
