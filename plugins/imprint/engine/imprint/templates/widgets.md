---
# Every widget Imprint draws, and the YAML behind it. Build it with:
#   imprint pptx widgets.md --pdf
# `imprint widgets` lists them; `imprint widgets fields` lists the standard fields.
TITLE: "Widgets"
SUBTITLE: "Charts and infographics from a YAML block"
DATE: "2026-10-01"
---

# Charts

PowerPoint charts: the reader can click in and edit the data.

## September was the best month since launch

```widget
widget: column
items:
  Apr: 64
  May: 88
  Jun: 102
  Jul: 120
  Aug: 180
  Sep: 240
highlight: Sep
note: Applications per month, app and branch
```

## The app is now the fastest way to open an account

- Five working days at a branch
- One day in the app
- The call centre sits between them

```widget
widget: bar
items:
  Branch: 5
  Call centre: 3
  App: 1
highlight: App
note: Working days from application to an open account
```

## Growth this year is ahead of last year every month

```widget
widget: line
categories: [Jan, Feb, Mar, Apr, May, Jun]
series:
  2025: [12, 14, 13, 15, 16, 18]
  2026: [14, 17, 19, 22, 26, 31]
highlight: "2026"
suffix: K
note: New accounts, thousands
```

## The app took over from the branch in three quarters

```widget
widget: stacked-column
categories: [Q1, Q2, Q3]
series:
  App: [40, 70, 110]
  Branch: [80, 75, 60]
```

## Most applications now start in the app

```widget
widget: doughnut
title: Where applications come from
suffix: "%"
items:
  App: 58
  Branch: 30
  Partners: 12
```

```widget
widget: pie
title: Customer segments
suffix: "%"
items:
  Retail: 64
  SME: 26
  Corporate: 10
```

## New clients more than covered this year's losses

```widget
widget: waterfall
prefix: "THB "
suffix: M
items:
  - {label: 2025, value: 120, total: true}
  - {label: New clients, value: 45}
  - {label: Price change, value: -12}
  - {label: Lost clients, value: -8}
  - {label: 2026, value: 145, total: true}
note: Revenue, million baht
```

# Numbers

Headline figures, progress and shares, drawn from shapes.

## The pilot moved every number that matters

```widget
widget: kpi
items:
  - {icon: user-plus, label: Accounts opened, value: "12,400", delta: +18%, note: vs 2025}
  - {icon: clock, label: Days to open, value: 1, delta: -4, note: was 5}
  - {icon: star, label: Customer rating, value: 4.6, delta: +0.8, note: out of 5}
```

## Build is on track; testing starts next month

```widget
widget: progress
suffix: "%"
items:
  - {icon: clipboard-list, label: Requirements, value: 100}
  - {icon: code, label: Build, value: 72, note: 18 of 25 screens}
  - {icon: flask-conical, label: Testing, value: 35}
  - {icon: rocket, label: Rollout, value: 0}
```

## Four in five applications need no one to touch them

```widget
widget: rings
suffix: "%"
items:
  - {label: Straight through, value: 82, note: no staff review}
  - {label: Returned, value: 11, note: missing documents}
  - {label: Rejected, value: 7, note: failed screening}
```

## Six in ten customers now apply in the app

- The app share doubled in a year
- Branches keep the customers who need help

```widget
widget: waffle
suffix: "%"
items:
  - {label: Apply in the app, value: 58}
```

# Flows and structure

Stages, milestones, cycles and the parts of a whole.

## One in six visitors opens an account

```widget
widget: funnel
items:
  - {label: Visited, value: 12000}
  - {label: Started, value: 4200, note: 35% of visitors}
  - {label: Submitted, value: 2600, note: 62% of starts}
  - {label: Opened, value: 2100, note: 81% of submissions}
```

## The pilot runs this quarter; rollout follows

```widget
widget: timeline
items:
  - {when: Q1, icon: search, label: Discovery, note: Interviews and data, status: done}
  - {when: Q2, icon: code, label: Build, note: App and staff screen, status: done}
  - {when: Q3, icon: flag, label: Pilot, note: Two branches, status: now}
  - {when: Q4, icon: rocket, label: Rollout, note: Every branch, status: next}
```

## Strategy is a loop, not a plan on a shelf

```widget
widget: cycle
center: Strategy
items:
  - {icon: radar, label: Scan, note: Markets and risks}
  - {icon: lightbulb, label: Formulate, note: Goals and choices}
  - {icon: hammer, label: Implement, note: Plans and budget}
  - {icon: chart-line, label: Evaluate, note: Measure and adjust}
```

## One platform serves every team

<!-- tone: dark -->

```widget
widget: hub
center: Platform
items:
  - {icon: user-plus, label: Onboarding}
  - {icon: scan-face, label: Identity}
  - {icon: shield-check, label: Screening}
  - {icon: landmark, label: Accounts}
  - {icon: file-chart-column, label: Reports}
```

## The app does what customers asked for most

```widget
widget: features
items:
  - {icon: zap, label: Open in minutes, note: "One form, no branch visit"}
  - {icon: scan-face, label: Verify with your face, note: Checked against the national ID service}
  - {icon: shield-check, label: Safe by design, note: Bank-grade encryption end to end}
  - {icon: wallet, label: No monthly fee, note: Free for personal accounts}
  - {icon: bell, label: Instant alerts, note: Every payment as it happens}
  - {icon: headset, label: Help when needed, note: A person answers in two minutes}
```

## Three screens carry the whole journey

```widget
widget: devices
frame: iphone
items:
  - {image: showcase-screen-1.png, label: Apply, note: One form}
  - {image: showcase-screen-2.png, label: Verify, note: Face and ID}
  - {image: showcase-screen-3.png, label: Track, note: Every step live}
```

## A two billion baht share is within reach

```widget
widget: nested
prefix: "THB "
suffix: B
items:
  - {label: Total market, value: 48, note: Retail deposits in the region}
  - {label: Served market, value: 12, note: Customers who bank digitally}
  - {label: Target share, value: 2.4, note: 20% of the served market by 2028}
```

## Start with the quick wins

```widget
widget: matrix
x: [Low effort, High effort]
y: Value
quadrants: [Quick wins, Big bets, Fill-ins, Money pits]
items:
  - {label: e-KYC, x: 25, y: 85}
  - {label: Core upgrade, x: 80, y: 78}
  - {label: New branding, x: 30, y: 30}
  - {label: Branch redesign, x: 78, y: 22}
```
