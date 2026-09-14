# Widgets

A widget is a chart or an infographic written as a YAML block. The engine draws it from
PowerPoint's own parts, so the reader can click into it and edit it:

- **Native charts** (column, bar, stacked-column, stacked-bar, line, area, pie, doughnut,
  waterfall) are real PowerPoint charts. Right-click, Edit Data, and the numbers open in Excel.
- **Infographics** (kpi, progress, rings, funnel, timeline, cycle, hub, nested, waffle,
  matrix) are drawn from PowerPoint shapes. Every bar, ring and label is a shape you can move.

A widget sits where a picture would: alone on a slide it takes the whole area, beside text it
takes the right side, and a wide one (a timeline, a line chart, three or more kpi tiles) goes
under the text at full width. Two widget blocks on one slide sit side by side.

In a Word document a widget becomes a table of its data, under its title, with its note.

```bash
imprint widgets              # the list
imprint widgets kpi          # an example block to copy
imprint widgets fields       # the standard fields
imprint new widgets deck.md  # a deck that uses every widget
```

## The standard

Every widget reads the same fields. A field a widget does not use is ignored; a field that is
not in the standard stops the widget with a warning that names it, so a typo never passes
silently.

````markdown
```widget
widget: column                  # which widget (required)
title: Applications per month   # optional, a short heading above it
items:                          # the data
  Jul: 120
  Aug: 180
  Sep: 240
highlight: Sep                  # optional, pick one out; the rest are muted
note: Branch and app channels   # optional, one line under it: source, period
```
````

| Field | Meaning |
| :-- | :-- |
| `widget` | the widget's name |
| `title` | a short heading above the widget; the slide title still states the point |
| `note` | one line under the widget: the source, the period, a footnote |
| `items` | the data: a list of items, or a `label: value` mapping for the simple case |
| `categories`, `series` | a chart with several series: the labels along the axis, and `name: [values]` per series |
| `prefix`, `suffix` | text around every number: `prefix: "THB "`, `suffix: "%"` |
| `decimals` | digits after the point; by default, as written |
| `max` | what a full bar, ring or waffle stands for; 100 when the values are percentages |
| `highlight` | a label or a list of labels to pick out; everything else is muted |
| `colors` | `one` (one colour for every item) or `each` (the theme's palette in order) |
| `labels` | `true` or `false`: the value on each bar, slice or point |
| `legend` | `auto`, `bottom`, `right` or `none`; auto shows one only for several series or a pie |
| `columns` | how many tiles, rings or cards sit side by side |
| `center` | the text in the middle of a hub or cycle |
| `x`, `y`, `quadrants` | a matrix's axes (`[low end, high end]` or a name) and its four quadrant names |

Inside `items`, each item takes:

| Item field | Meaning |
| :-- | :-- |
| `label` | what the item is |
| `value` | its number; a kpi also takes short text such as `"5 to 1"` |
| `note` | one short line about it |
| `delta` | the change beside a kpi: `+18%`, `-4` (a rise is green, a fall red, with a small triangle shape) |
| `max` | this item's own full scale |
| `when`, `status` | a timeline's date or period, and `done`, `now` or `next` |
| `total` | `true` for a waterfall bar that shows a running total |
| `x`, `y` | a matrix item's position, 0 to 100 |

Quote a value YAML would misread: `"2026"` as a label that must stay text, `"%"`, `"12,400"`,
and any text with a colon.

## Choosing a widget

| The point of the slide | Widget |
| :-- | :-- |
| One to four headline numbers and how they moved | `kpi` |
| Compare a few amounts; one of them is the story | `column` with `highlight` (`bar` when labels are long or it is a ranking) |
| A trend over time | `line`; `area` when the volume matters |
| How parts make up a whole, at most six parts | `doughnut` or `pie`; `waffle` for a single share |
| The parts changed over periods | `stacked-column` |
| How a total got from one figure to another | `waterfall` |
| Progress against a target or plan | `progress`; `rings` for two to four percentages that stand alone |
| People dropping out at each stage | `funnel` |
| Where a project stands | `timeline` with `status` |
| A process that repeats | `cycle` |
| A platform or organisation and its parts | `hub` |
| A market and the share you can win | `nested` |
| Priorities on two dimensions | `matrix` |

When none fits, draw the diagram with the diagram-design skill in the `slide-16x9` preset
and place the PNG, or build a table.

## How good data slides look

These come from how professional pitch, data and strategy decks are built:

- **The title says what the data shows**: "September was the best month since launch", not
  "Applications per month". The widget's own `title` is for the measure, and is often not needed.
- **Pick out one thing.** Mute everything else with `highlight`. A chart where every bar is a
  different colour says nothing.
- **Label directly.** Values sit on the bars, so the value axis and gridlines go; the engine
  does this whenever labels are on.
- **Round to what the audience needs**: 12.4K, not 12,437. Say what a number is compared
  with, in `delta` or `note`.
- **One widget, one message.** Two widgets on a slide only when they are read together (a
  channel doughnut beside a segment pie). Beside text, keep the text to three short lines.
- **Big numbers are big.** A kpi slide has at most four tiles; more is a table.
- **State the source** in `note` for any figure a client could question.
- **Colour means the same thing on every slide.** The theme's palette is applied in order; keep
  items in the same order across slides so a colour always names the same thing.

## Limits the build warns about

Column and stacked charts beyond twelve categories, pie and doughnut beyond six parts, kpi
beyond eight tiles, rings beyond five, funnel beyond six stages, timeline beyond seven
milestones, cycle and hub beyond eight. The widget still draws; the slide should be split.

## Theme keys

A theme's `slides.widgets` section sets the colours; `slides.size.widget_value` and
`widget_label` set the type.

| Key | What it sets |
| :-- | :-- |
| `palette` | series and item colours, in order (default: `components.accents`) |
| `highlight` | the colour of what `highlight` picks out, and of one-colour widgets |
| `muted` | everything a highlight leaves out |
| `track` | the empty part of a bar, ring or waffle |
| `grid` | axis lines and gridlines |
| `positive`, `negative` | rises and falls, in a kpi and a waterfall |
| `dark` | the same keys for `tone: dark` slides |
