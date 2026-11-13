# Dashboard design system

This document defines the visual foundation for the redesigned Dashboard. The
current implementation is intentionally scoped to `pages/1_Dashboard.py`; the
remaining pages continue to use the shared application theme until they are
redesigned.

## Design direction

The Dashboard uses a compact operations-workspace style: deep navy provides the
primary structure, crimson marks active or selected information, and white data
surfaces keep dense reporting content readable. Spacing is deliberately tight so
charts and tables feel related and the page does not develop empty visual zones.

## Color tokens

| Token | Value | Use |
| --- | --- | --- |
| Navy 950 | `#081A36` | Dashboard header gradient and deepest emphasis |
| Navy 900 | `#091D3C` | Primary metric gradient |
| Navy 800 | `#10294F` | Headings and primary text |
| Navy 700 | `#173B6D` | Charts, controls, and navy accents |
| Navy 500 | `#31598D` | Secondary chart series |
| Crimson 800 | `#7C102A` | Crimson gradient shadow |
| Crimson 700 | `#A81737` | Crimson borders and emphasis |
| Crimson 600 | `#B6193B` | Primary selection/action color |
| Crimson 500 | `#C32148` | Crimson gradient highlight |
| Crimson 400 | `#D84A67` | Secondary chart series |
| Ink 600 | `#52647D` | Labels and secondary text |
| Ink 500 | `#687A92` | Chart axes and metadata |
| Canvas | `#F5F7FA` | Application background |
| Surface | `#FFFFFF` | Cards, panels, and tables |
| Border | `#DDE4EE` | Panel and control boundaries |
| Grid | `#E9EEF5` | Chart grid lines |

### Gradients

- Title accent: `linear-gradient(180deg, #173B6D, #B6193B)`
- Navy metric: `linear-gradient(138deg, #091D3C, #173B6D)`
- Crimson metric: `linear-gradient(138deg, #7C102A, #C32148)`
- Selected supplier chip: `linear-gradient(112deg, #142F57, #A6193A)`

## Typography

The interface uses **Inter** when available, followed by the operating system's
sans-serif UI font. Streamlit already ships with Inter, so no remote font request
is required.

The Dashboard follows a fixed type scale. Do not introduce one-off sizes.

| Role | Size | Weight | Line height |
| --- | --- | --- | --- |
| Product label | `11px / .7rem` | 750 | 1.2 |
| Dashboard title | `25px / 1.55rem` | 750 | 1.1 |
| KPI value | `30px / 1.875rem` | 750 | 1.0 |
| Panel title | `15px / .96rem` | 700 | 1.25 |
| Body/control text | `14px / .875rem` | 400–600 | 1.4 |
| Label/caption | `11–12px / .69–.75rem` | 600–700 | 1.2–1.3 |
| Chart label | `11px` | 400 | default |

Numeric values use tabular figures. Headings use a slightly negative letter
spacing; uppercase labels use positive tracking. On small screens only the title
and KPI values step down, preserving the same hierarchy without wrapping.

## Spacing and shape

- Base internal spacing: `10px` (`.625rem`).
- Panel padding: `13–14px`.
- Control height: `42px`.
- Panel and KPI radius: `14px`.
- Mobile metric grid: two columns, never a single long stack.
- Use small gaps between related chart/table panels; avoid decorative whitespace.

## Component behavior

- Fiscal year and quarter remain adjacent in the filter bar.
- Packaging suppliers use a multi-select dropdown. Selected suppliers appear as
  navy-to-crimson gradient chips and remain removable from the control.
- The page title is an unboxed, compact two-line lockup: `Matalan Carton Supply
  Report` followed by `Dashboard`.
- All four KPI cards are clickable and keep a stable visual identity in every
  state: Total is navy, Packaging suppliers is crimson, and Standard/Custom use
  white data surfaces. A raised outline identifies the active card without
  changing its base color, size, or typography.
- Card selection uses an in-page Streamlit component inside the Dashboard
  fragment. It updates the cards, charts, and tables without navigating or
  reloading the full application.
- `Total cartons` shows the full active filter selection.
- `Packaging suppliers` shows the number of selected suppliers and switches the
  detail table to aggregated supplier totals.
- `Standard cartons` and `Custom cartons` filter both charts and the Live Summary
  table to that carton type. `Total cartons` restores the complete filtered view.
- Tracker/eBiz reconciliation responds to supplier selection but represents the
  complete reconciliation dataset because that source is not quarter-scoped.
- Tables size to their content up to a capped height, then scroll internally.

## Accessibility

- White text is used only on the darkest navy/crimson gradient areas.
- Data is never differentiated by color alone; labels, legends, and values remain
  visible.
- Controls keep explicit labels and keyboard behavior supplied by Streamlit.
- Motion follows the existing `prefers-reduced-motion` rule.
- Card and data-panel changes use a short 200ms fade/slide transition; the
  transition is disabled when the visitor requests reduced motion.

## Build reports extension

The Build reports page reuses the Dashboard's typography, crimson/navy palette,
compact spacing, 14px panels, gradient supplier chips, and responsive behavior.

- The old PACD/Matalan eyebrow and descriptive tagline are not rendered.
- A compact two-line title is paired with a navy-to-crimson Available suppliers
  count badge.
- Packaging supplier selection occupies the first form row. Fiscal year, quarter
  selection share the second row and stack naturally on narrow screens. The
  discrepancy option sits in the Report selection panel's top-right header area.
- Selected quarters use compact, content-sized pills with a restrained navy
  active state.
- Build, Download all, and Save form one compact action sequence in that order.
  Download remains disabled until the current selection has been built, and the
  sequence sits at the right side of the Individual files panel header. On mobile,
  all three actions remain on one line as 36px square icon-only buttons.
- Build uses a compact crimson gradient, Download uses navy, and Save uses a
  neutral white secondary action. None stretches to fill the page. Individual filenames use compact
  neutral white controls with navy-gray text; crimson is not used for filenames.
  Filename labels are left-aligned at every breakpoint and use a smaller type size
  on mobile.
- Report generation, caching, local saving, discrepancy selection, and individual
  downloads retain their existing behavior.

All pages inherit one global responsive content inset from `stMainBlockContainer`:
the same top, left, right, and bottom values are applied regardless of which page
is active. Only viewport breakpoints alter those shared values.

## Cross reference extension

- The application requests an expanded sidebar on desktop. Streamlit retains its
  responsive collapsed treatment on mobile screens.
- Cross reference uses the compact two-line page heading with an eBiz row-count
  badge and a single supplier filter surface.
- Six reconciliation cards reuse the Dashboard component and rerun only the
  Cross reference fragment when selected. They display six columns on wide
  screens and two columns at responsive breakpoints. Alternating navy and
  crimson gradients distinguish the six categories while a crimson focus ring
  identifies the active view.
- Card state is persistent: one click updates the active outline and filters the
  appropriate Tracker or eBiz-only table. Existing card nodes are retained on
  fragment reruns so selection changes do not flash or replay a grid animation.
- The eBiz badge always displays a numeric row count, including `0` when no eBiz
  file has been loaded.
- Tracker and eBiz-only results use compact white panels, content-sized tables,
  searchable rows, and a navy workbook download action.
- Matching, filtering, cache behavior, table values, and workbook generation are
  unchanged.
