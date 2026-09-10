# Writing standards for every document

The rules every Imprint document follows, whatever its type. They rest on published
standards where one exists; where none does, on common professional practice.

## Plain language

Follow ISO 24495-1:2023, *Plain language, Part 1: Governing principles and guidelines*. A
document is plain when its intended readers can find what they need, understand it, and use
it. In practice:

- Put the conclusion first. The first paragraph of a section says what the section decides
  or asks; the detail follows.
- One idea per paragraph, one requirement per sentence.
- Short sentences, active voice, the actor named: "The system sends a receipt", not
  "A receipt is sent".
- Define every abbreviation on first use and list them in a glossary.
- Use the reader's words for their business, not the team's internal names.

## Words that bind

In a specification, the verb carries the obligation. Use one convention per document and
say which in its introduction:

| Word | Meaning | Source |
| :-- | :-- | :-- |
| shall | a requirement: binding, verifiable | ISO/IEC/IEEE 29148 |
| should | a recommendation: expected, but a reason can justify not doing it | ISO/IEC/IEEE 29148 |
| may | permitted, optional | ISO/IEC/IEEE 29148 |
| will | a statement of fact or intent, not a requirement | ISO/IEC/IEEE 29148 |
| MUST, SHOULD, MAY | the same levels, in capitals, for protocol and API documents | RFC 2119, RFC 8174 |

Never use "must" and "shall" for the same thing in one document.

## Numbers, dates and units

- **Dates in prose**: spelled out, "1 October 2026". **Dates in tables, IDs and data**:
  ISO 8601, `2026-10-01`. Never `01/10/26`, which reads differently in different countries.
- **Times**: 24-hour, with the time zone when more than one is involved: `14:30 ICT (UTC+7)`.
- **Units**: SI units and symbols (ISO 80000), a space between number and unit: `16 GB`,
  `200 ms`.
- **Money**: the currency in the column header, not in every cell; say whether VAT or other
  tax is included.
- **Figures in a column**: the same number of decimals down the column, a thousands
  separator, right-aligned (`--:`).

## Tables

Align each column by what it holds: text left, quantities and money right, short values of
equal length (Yes/No) centred, codes and IDs left. The header follows its column. A table
has a heading or an introducing sentence; a reader should know what it is before reading it.

## Figures

Every figure has a caption, written as the image's alt text: `![Figure 3: Order flow](file.png)`.
The caption says what the figure shows, not "Diagram". Refer to it by number in the text.
Draw diagrams at document size so their smallest label prints at 6 pt or more.

## Document control

Every controlled document carries, straight after the cover (ISO 9001:2015 clause 7.5,
documented information):

- **Identification**: title, document ID or reference, version, date, status
- **Version history**: version, date, author, description of the change
- **Review and approval**: role, name, date, and signature where the client needs one

Versions: `0.x` while drafting, `1.0` on first approval, `1.1` for a correction, `2.0` for a
change of substance. Status is one of Draft, For Review, Approved, Superseded.

For which information a document of each kind should contain, ISO/IEC/IEEE 15289:2019
(*Content of life-cycle information items*) is the general reference.

## Accessibility

- Alt text on every picture (the caption does this).
- Text contrast of at least 4.5:1 against its background (WCAG 2.2, criterion 1.4.3).
- Never carry meaning by colour alone; add a label, a pattern or a symbol drawn in the figure.
- Headings are real headings (`#`, `##`), not bold lines, so the document has a navigable outline.

## Plain text

No section signs, emoji or decorative symbols in headings, body, tables or captions: arrows,
check marks, crosses, stars, typed bullets, em-dashes used as punctuation. Write "Section 3",
"to", "Yes". Many readers do not know what a section sign means, and many fonts do not carry
these glyphs, so Word substitutes another font and the line looks broken.
