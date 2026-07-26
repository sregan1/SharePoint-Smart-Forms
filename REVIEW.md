# Smart Forms — Full Application Review

> **Status: implemented.** Everything in phases 1–5 below has been built, plus the
> templates gallery and save-and-resume from the "also worth considering" list.
> The build is green (tsc, eslint, webpack) with 107 unit tests passing.
>
> **Deliberately not shipped**, with reasons:
> - **QR code in the share dialog.** A correct QR encoder needs Reed–Solomon
>   error correction, and I had no way to verify the output was scannable without
>   a decoder. A wrong QR code is worse than no QR code, so it was dropped rather
>   than guessed at.
> - **Excel `.xlsx` and PDF export.** Both need a third-party library; the CSV
>   export covers the same need without adding weight to an SPFx bundle.
> - **Page-level branching (skip-to-section).** Per-field and per-section
>   visibility rules cover the common cases; a full jump graph needs cycle
>   detection and its own editor, which is a feature in its own right.
> - **Address as separate columns.** Stored as one formatted line instead — six
>   columns per question would make the list view and every export unmanageable,
>   and address parts are never analysed independently. Noted in the README.
>
> One gap worth knowing about: `sanitizeHtml`'s DOM tree-walk can't be unit tested
> without a DOM, so the tests cover the allowlist tables and the URL-scheme
> decision directly and assert the no-DOM fallback is safe. The tree walk itself is
> exercised only at runtime.

Reviewed at ~5,400 lines across 22 source files. SPFx 1.20 / Fluent UI 8 / React 17 / PnPjs 3.26.

**Overall:** the architecture is genuinely good. The WYSIWYG designer (the canvas *is* the form) is the
right model, the CAML provisioning path is clean, and `formUtils.ts` is well-factored. What holds it back
from "ultimate forms app" is: one critical persistence trap, ~12 real bugs, a results experience that is
currently a table plus bar charts, and a field-type catalog that has breadth but is missing the two or
three types people actually ask for first.

---

## 1. Critical issues

### 1.1 Nothing tells the user their form isn't saved until the page is republished

The form definition lives in web part properties. `persist()` debounces 600 ms then calls
`onFormDefinitionChange`, which sets `this.properties.formDefinition` — but SPFx only writes properties to
the page when the user **saves/publishes the page**. There is no indicator anywhere.

Failure path: owner builds a 20-question form → clicks **Collect responses** → columns are created in
SharePoint → copies the share link → navigates away without republishing → **the entire form definition is
gone** while the list columns remain. Unrecoverable, and the app gave no warning.

Fix: a persistent status chip in the owner bar — `Draft — publish the page to save` / `All changes saved`.
On `handleCollectResponses`, block or warn until the page is published. Ideally also mirror the definition
into a hidden list item or the list's `RootFolder` property bag so it survives an unpublished page.

`src/webparts/smartForms/components/SmartForms.tsx:122-151`

### 1.2 "Preview" writes real responses to the list

The preview hint literally says *"answers submitted here are saved to the list."* A preview that pollutes
your response data isn't a preview. Add a `readOnly`/`dryRun` prop to `FormRenderer` that validates and
shows the confirmation screen without calling `submitResponse`.

`src/webparts/smartForms/components/SmartForms.tsx:342-345`

### 1.3 Rich text is stored unsanitized into a `RichTextMode="FullHtml"` column

`RichTextField` hands `editorRef.current.innerHTML` straight through. A respondent can paste
`<img src=x onerror=...>` or `<iframe>`; it lands in a FullHtml Note column and gets rendered by the OOB
list view and any downstream Power BI/Lists surface. Smart Forms' own display strips tags so *this* UI is
safe, but the stored value is a stored-XSS vector for every other consumer.

Fix: sanitize on `onPaste` and before `onChange` — allowlist `b i u strong em ul ol li p br a[href^=http]`,
strip every attribute not on the allowlist, strip all `on*`.

`src/webparts/smartForms/components/form/RichTextField.tsx:40-50`

### 1.4 Responses break past 5,000 items

`getResponses` does `.filter("SFStatus eq 'Complete'").top(2000)` on a column that is never indexed. Once
the list crosses the 5,000-item list view threshold this throws and the whole Responses tab shows the
generic "publish the form once" error — which is misleading and unactionable.

Fix: add `Indexed="TRUE"` to `buildStatusFieldXml`, page with `getPaged()`, and show a real "showing first
N of M" affordance instead of silently truncating at 2,000.

`src/webparts/smartForms/services/SharePointService.ts:29,356-360`, `src/webparts/smartForms/utils/spFieldXml.ts:110`

---

## 2. Confirmed bugs

| # | Bug | Location |
|---|---|---|
| 1 | **Changing a question's type leaves the old type's config behind.** `patchField({ type })` from the "Change type" menu doesn't seed the new type's defaults or drop the old ones. Text → Choice gives a `ChoiceGroup` with zero options; Choice → Rating keeps orphan `choices`; → Slider has no `min`/`max`/`step`. Reuse the seeding block from `addField`. | `FormDesigner.tsx:288` |
| 2 | **Deleting a section leaves dangling branching rules.** `removeField` cleans up `visibleWhen` referencing it; `removeSection` doesn't. `isFieldVisible` returns `true` when the driver is missing, so branched questions silently become permanently visible. | `FormDesigner.tsx:195-202` |
| 3 | **Duplicate option labels crash/misrender.** `key={c}` on `Checkbox`/`ChoiceGroup` options and `key={option}` in `RankingField`. The designer happily allows two options named "Other". React key collision → wrong option toggles. Use index-based or composite keys. | `FieldControl.tsx:207,222` · `RankingField.tsx:43` |
| 4 | **Sliders look answered but submit nothing.** The control renders `min` when `value` is `undefined`, so the user sees "0" and moves on. `toSharePointValue` returns `undefined` → the column stays empty. Either initialize the value to `min` on mount, or render an explicit unset state. | `FieldControl.tsx:294-307` |
| 5 | **Wizard can show a completely blank step.** If every field in a section is hidden by branching, `renderSection` returns `null` and the user gets an empty page with a Next button. Skip sections with no visible fields when computing steps. | `FormRenderer.tsx:264-268,282` |
| 6 | **Detail panel opens on focus, not on click.** `onActiveItemChanged` fires during keyboard navigation, so arrowing through the grid pops the panel repeatedly. Use `onItemInvoked`, or a dedicated chevron/"Open" column. | `ResponsesView.tsx:174` |
| 7 | **Ranking and multi-choice values break on semicolons.** Ranking persists as `"A; B; C"` and the summary re-splits on `;`; an option containing a semicolon corrupts the parse. Same for `defaultValue` splitting on multi-choice. Store ranking as JSON, or reject `;` in option text. | `SharePointService.ts:292-296` · `SummaryView.tsx:110-113` |
| 8 | **Responses reload on every re-render.** `load`'s `useCallback` depends on the `definition` object, and the `formDefinitionJson` effect creates a brand-new object on every web part render. Depend on `formDefinitionJson` (the string) instead. | `ResponsesView.tsx:56` · `SmartForms.tsx:153-155` |
| 9 | **Search filter goes stale.** The `filtered` memo reads `fields` but only depends on `[items, search]`. | `ResponsesView.tsx:60-72` |
| 10 | **Percent is a lie.** `buildFieldXml` writes `Percentage="FALSE"`, so a "45%" answer is stored as the number `45` and displayed as `45` in the list while the form shows `45%`. Also `formatResponseValue` ignores `decimalPlaces` for Percent. Either commit to `Percentage="TRUE"` with 0–1 storage, or drop the type (see §4). | `spFieldXml.ts:70-74` · `formUtils.ts:212-215` |
| 11 | **Currency is hardcoded to US.** `LCID="1033"` regardless of `currencySymbol`, so a "£" form renders "$" in the list view. Map symbol → LCID, or expose a locale picker. | `spFieldXml.ts:77` |
| 12 | **Only the first four questions are ever visible in the grid.** `MAX_GRID_FIELDS = 4`, no column picker, no indication that columns were dropped. | `ResponsesView.tsx:29,88` |
| 13 | **DateTime silently rounds.** A stored 09:15 shows "9:00 AM" in the dropdown because it floors to a 30-minute slot, and re-saving overwrites the real value. | `FieldControl.tsx:154` |
| 14 | **No `min`/`max` validation for Slider, Rating, or NPS**, only for Number/Currency/Percent. | `formUtils.ts:113-124` |
| 15 | **Dead draft/auto-save scaffolding.** `ResponseStatus = 'Draft' \| 'Complete'`, `IResponseItem.status`, `.saveIndicator`/`.saveIndicatorError` SCSS, and the `SFStatus` "used by auto-save" comment all exist — but there is no draft feature and `status` is never rendered. Either build save-and-resume (worth doing, §4) or delete it. | `models/index.ts:208-217` · `FormRenderer.module.scss:319-333` |
| 16 | **Dead props.** `title`, `context`, and `onListChange` are declared on `ISmartFormsProps`, wired in the web part, and never read. | `SmartForms.tsx:25-35` |
| 17 | **Items created directly in the list are invisible.** `SFStatus` defaults to `Draft`, and pre-existing items have no value at all, so the `eq 'Complete'` filter hides everything not submitted through Smart Forms. Reasonable intent, but nothing tells the owner why their 200 existing rows vanished. | `SharePointService.ts:359` |
| 18 | **Nothing warns about respondent permissions.** The share dialog hands out a link, but a respondent without Contribute on the list just gets a failed submit with a "make sure the form has been published" message that blames the wrong thing. Check `AddListItems` for the target audience, or at least explain it in the share dialog. | `SmartForms.tsx:364-385` |

---

## 3. UI/UX review

### What's already right
The pill-tab owner bar, the click-to-expand question card, Enter/Tab/Backspace option editing, the
accent-color CSS variable threading, and the `choices.length <= 6 ? buttons : dropdown` heuristic are all
good instincts. Don't undo these.

### 3.1 The whole UI is invisible in a themed or dark site
Every color is a hardcoded hex — `#323130`, `#ffffff`, `#faf9f8`, `#edebe9`. On a dark-themed SharePoint
site the form card renders white-on-white-ish and the designer becomes unreadable. This is the single
biggest polish gap.

Fix: thread Fluent theme tokens into CSS custom properties once at the root
(`--sf-bg`, `--sf-fg`, `--sf-fg-muted`, `--sf-border`, `--sf-card`, `--sf-hover`) from
`this.context.sdks`/`useTheme()`, and replace every literal. It's mechanical and it transforms how
professional the app feels.

### 3.2 The form title is buried
The most prominent thing on the rendered form — the title — is edited three clicks deep in a panel called
"Style". Microsoft Forms puts it at the top of the canvas. Make the title and description directly editable
as the first block of the designer canvas, and drop them from the panel.

### 3.3 Rename and re-split the settings panel
"Style" currently contains the title, description, color, wizard toggle, thank-you message, **email
notifications**, and an "Advanced options" accordion. Notifications are not style. Split into a single
**Form settings** panel with clear groups: *Basics · Appearance · After submit · Notifications · Access*.

### 3.4 No question reordering by drag
Up/down `IconButton`s only. Ironic, given `RankingField` already implements HTML5 drag-and-drop — lift that
into a small reusable `useDragList` hook and use it for questions and sections too.

### 3.5 Long forms have no navigation
No collapse-all, no question search, no outline/minimap, no jump-to-section. At 40 questions the designer is
an unmanageable scroll. Add a collapsible left rail listing questions (number + truncated title + branch
icon) that scroll-syncs and doubles as the drag target for reordering.

### 3.6 "Collect responses" doesn't validate the form first
It happily provisions columns for a form with untitled questions (auto-named `Question 7`), Choice fields
with one blank option, or branching rules pointing at deleted fields. Run a pre-flight check and show a
dismissible list of warnings with "jump to question" links.

### 3.7 Empty and first-run states are thin
The `!listId` placeholder is good. But: the designer's zero-question state is just an "Add new" button, the
Responses tab's empty state is italic gray text, and there's no template gallery. See §5.1.

### 3.8 Field-level polish
- `<label className={styles.fieldLabel}>` has no `htmlFor` and wraps nothing — screen readers never
  associate the label with the input. No `aria-required`, `aria-invalid`, or `aria-describedby` on any
  control. This is a compliance problem for a form product, not a nice-to-have.
- The NPS row has `role="radiogroup"` / `role="radio"` but no arrow-key navigation and no `aria-label`.
- Half-width fields are the only layout option — no thirds, no explicit row grouping.
- `field.width` exists in the model but **there is no UI to set it**. Ship the control or remove the field.

---

## 4. Field types — keep, add, remove

### 4.1 Remove or fold in (3)

| Type | Why |
|---|---|
| **Location** | It is a `TextField` with a map-pin icon and a `MaxLength`. It provides nothing a Text field with a placeholder doesn't. Either delete it, or make it real: Bing/Azure Maps autocomplete storing a structured address + lat/long. Right now it's catalog padding. |
| **Percent** | Broken (§2.10) and redundant with Number. Fold into Number as a "display as percentage" option. |
| **Currency** | Same argument — fold into Number as a "format: number / currency / percent" dropdown with a symbol + locale. Three separate types for one concept bloats the "Add new" menu without adding capability. |

Also fold **Date** and **DateTime** into one **Date** type with an "include time" toggle — same CAML type,
same control, one fewer entry in the picker.

Net: the picker drops from 19 entries to 14 while *losing no capability*, which directly serves the
"uncluttered" goal.

### 4.2 Add — the must-haves (5)

1. **File upload** — the #1 missing type, no contest. Microsoft Forms has it; SharePoint lists support
   attachments natively (`item.attachmentFiles.add()`). Owner sets allowed extensions, max size, and
   max count. Without this the app can't handle expense claims, applications, or anything with a document.
2. **Likert / matrix grid** — rows of statements × a shared column scale ("Strongly disagree → Strongly
   agree"). Every survey needs it. Store as one Note column of JSON, or one Choice column per row.
3. **Signature** — canvas capture stored as a base64 PNG attachment. Unlocks approvals, waivers, sign-offs,
   H&S forms. High perceived value, ~120 lines.
4. **Lookup from another list** — choices sourced live from a SharePoint list column instead of a static
   array. This is the thing Smart Forms can do that Microsoft Forms structurally cannot, and it's the
   strongest reason to pick this web part. Cache the options; support a filter and a display column.
5. **Section break / static content block** — a non-input block for instructions, an image, a divider, or
   a rich-text explanation between questions. Currently the only way to add guidance is a section
   description, which forces an artificial section.

### 4.3 Add — high value (5)

6. **Time only** — a `TimeOnly`-style picker; currently you must use DateTime and ignore the date.
7. **Image choice** — options as picture tiles. Big polish win for product/design surveys.
8. **Consent / acknowledgement** — a single required checkbox with rich-text terms. Legally distinct from
   Yes/No and semantically clearer in results ("98% consented" vs. "98% Yes").
9. **Address (structured)** — street / city / state / postcode / country as one logical question storing
   separate columns. Requested constantly; the current answer is five Text fields.
10. **Numeric scale with custom labels** — generalize `NpsField` so the range, step, and endpoint labels are
    configurable (1–5 satisfaction, 1–7 agreement). NPS then becomes a *preset* of this, not a special case.

### 4.4 Rework existing

- **Yes/No**: "required means it must be Yes" is surprising. Split into `Toggle` (never required, always
  stores true/false) and use the new **Consent** type for must-be-checked semantics.
- **Rating**: allow half-stars and let the icon be a star / heart / thumb.
- **Choice**: add an **"Other" option with a write-in box** — the most-requested Choice feature there is.
  Also add option shuffling for survey rigor, and a min/max-selections rule for multi-select.
- **Person**: support group/SharePoint-group selection, not just `PrincipalType: 1` users.
- **Number**: add a units suffix (kg, hrs) now that Currency/Percent are folded in.

### 4.5 Cross-cutting field features worth adding

- **Regex / pattern validation** with a custom error message (postcodes, employee IDs, NHS numbers).
- **Custom required-message per field** — "Name is required" is fine; "We need this to contact you" is
  better.
- **Read-only / prefilled-and-locked fields** — pairs with the existing (and excellent, undersold) URL
  prefill feature.
- **Calculated / derived field** — e.g. `Total = Qty × UnitPrice`, evaluated client-side.
- **Multi-condition branching.** `visibleWhen` supports exactly one rule. Real forms need
  AND/OR groups, plus numeric/date comparison operators (`>`, `<`, `between`, `before`, `after`) — the
  current operator set is string-only. This is the single highest-leverage upgrade to the logic engine.
- **Page-level branching** (skip to section) for the wizard layout, not just per-field visibility.

---

## 5. Results & dashboard

Right now: a `DetailsList` capped at four columns, free-text search, CSV export, and a summary grid of
horizontal bars. That's a competent *summary*. It is not a dashboard, and this is where the biggest
opportunity is.

### 5.1 Proposed structure — four views behind the existing Pivot

```
Responses    ▸  [ Dashboard ]  [ Summary ]  [ Table ]  [ Individual ]
```

**Dashboard** (new, becomes the default)
- **KPI strip** — Total responses · Completion rate · Median time to complete · Responses today/7d/30d ·
  Unique respondents, each with a sparkline and a vs-previous-period delta.
- **Response volume over time** — an area chart with day/week/month granularity. Requires only `Created`,
  which is already loaded. Highest value per line of code in this entire proposal.
- **Auto-selected highlight cards** — pick the 4–6 most "interesting" questions automatically (an NPS
  gauge, the choice question with the most lopsided distribution, the rating with the lowest average, the
  numeric field with the widest spread) so the dashboard is useful with zero configuration.
- **Configurable tile layout** — let the owner pin/unpin/reorder tiles per question and pick the chart form
  (donut / bar / column / stat / gauge / word cloud). Persist the layout in `IFormSettings` so
  "options for any occasion" is real rather than a fixed template.

**Summary** (evolve the existing view)
- Chart-type switcher per card — bar, donut, column, stacked, table.
- Real numeric distribution: histogram + median and standard deviation, not just min/avg/max.
- Text questions: word/phrase frequency cloud + "show all N answers" drill-in, instead of three
  truncated samples.
- **Cross-tab / segment-by.** Pick any Choice/YesNo/Person question as a segment and every other card
  re-renders split by it ("NPS by department"). This is the feature that makes a forms product feel
  genuinely analytical, and nothing in Microsoft Forms does it.
- Click any bar segment → filters the Table view to those respondents. Charts that don't drill through are
  decoration.

**Table** (fix the current grid)
- Column picker with show/hide/reorder, persisted.
- Sortable columns (`onColumnClick` — currently absent), plus per-column filters.
- Sticky header, virtualized rows (`DetailsList` already supports this), frozen first column.
- Row selection with bulk delete/export, and inline editing of a response.
- Saved views ("This month", "Detractors only") stored in settings.

**Individual** (evolve `ResponseDetailPanel`)
- Prev/next navigation between responses without closing the panel.
- Render the response as the *form itself*, read-only — far more legible than a label/value list.
- Print / PDF-friendly stylesheet.
- Render rich text as sanitized HTML instead of stripping to plain text.

### 5.2 Cross-cutting results features

- **Global date-range and segment filter bar** applying to all four views at once.
- **Export**: Excel `.xlsx` with formatting alongside the current CSV; PDF of the dashboard; "open in
  Power BI" deep link. The CSV export already correctly emits a UTF-8 BOM — good detail, keep it.
- **Response detail: show `SFStatus`** — the column and the model field already exist.
- **Empty states with intent** — "No responses yet" should offer *Copy share link* and *Preview the form*,
  not italic gray text.
- **Charts must follow the accent color and be colorblind-safe.** A single `--sf-accent` bar fill works for
  one series; the moment you add donuts and segments you need a validated categorical palette. The
  `dataviz` skill in this environment covers exactly this — use it when building the charts.
- Chart rendering: build inline SVG components rather than adding a chart library. The shapes needed
  (bar, column, donut, area, gauge, sparkline, histogram) are ~400 lines total and keep the bundle small,
  which matters for SPFx.

---

## 6. Performance & optimization

| Issue | Detail |
|---|---|
| **O(n²) question numbering** | `questionNumber()` and `numberOf()` call `allFields(definition)` *inside* a per-field render, rebuilding the whole array for every question. 50 questions ≈ 2,500 array constructions per render. Memoize an `id → index` map once per render. `FormDesigner.tsx:303` · `FormRenderer.tsx:238` |
| **Full deep clone per keystroke** | Every designer edit runs `JSON.parse(JSON.stringify(definition))` — on each character typed into a title or option. Use targeted immutable updates, or debounce the clone. `FormDesigner.tsx:34,63` |
| **Serial field provisioning** | `ensureFields` awaits `createFieldAsXml` **and** `defaultView.fields.add` sequentially per field — 2 round trips × N fields. A 30-question form is 60+ serialized requests, easily 30–60 s. Use a PnPjs `[batch]`, or at minimum parallelize with a concurrency cap of 4. `SharePointService.ts:98-110` |
| **Serial `ensureUser`** | One round trip per person on submit. Batch them. `SharePointService.ts:309-313` |
| **`TIME_OPTIONS` built at module load** | 48 objects, trivial, but it runs even for forms with no DateTime field. Lazy-init. |
| **Whole-list refetch on delete** | Already handled optimistically via `setItems` — good. But `Refresh` refetches all 2,000 rows; add `getPaged`. |
| **No memoization on `SummaryView`** | Every distribution/NPS/ranking stat recomputes on any parent re-render across up to 2,000 items × N fields. Wrap in `useMemo` keyed on `items.length` + definition version. |
| **Bundle** | Fluent 8 `@fluentui/react` barrel imports pull a lot. Check `release/component-dependency-audit/` and switch hot paths to specific-component imports. |

---

## 7. Code health

- **`isFillView()` / `buildShareUrl()` read `window.location` directly** — breaks in the SPFx Workbench and
  makes the share flow untestable. Route through `context.pageContext`.
- **Two near-identical formatters.** `formatResponseValue` (SharePoint shape) and `formatFormValue` (form
  shape) duplicate ~80% of their switch bodies. Normalize to one internal shape and keep a single
  formatter.
- **`Math.random()`-based `newId()`** — fine in practice, but `crypto.randomUUID()` (or a counter + time) is
  free and removes the collision question entirely.
- **`generateDeterministicGuid`** uses a paired 32-bit FNV hash. Unique-per-list in practice since internal
  names are unique, but it's the kind of thing that fails once, mysteriously, in production. A comment
  explaining the uniqueness invariant would earn its keep.
- **`optionRefs` is a single flat array shared across all fields** and never cleared when the active
  question changes. Works because only one card is expanded at a time — fragile. Key it by field id.
- **No `schemaVersion` migration path.** `schemaVersion: 1` is written and never read. Add a
  `migrate(definition)` step now, before v2 exists and you need it retroactively.
- **Zero tests.** For a project whose core is pure functions — `validateField`, `isFieldVisible`,
  `generateInternalName`, `buildFieldXml`, `buildCsv`, `formatResponseValue`, `npsStats`, `rankingStats` —
  this is very cheap to fix and would have caught at least four of the bugs in §2.
- **`FormDesigner.tsx` is 557 lines** doing state, mutation helpers, option keyboard handling, menu
  construction, and three render paths. Split out `QuestionCard`, `SectionHeader`, and a `useFormEditor`
  hook.

---

## 8. Suggested order of work

**Phase 1 — trust (do these first, nothing else matters if the app loses data)**
1. Unsaved-changes indicator + publish-the-page warning (§1.1)
2. Non-destructive preview (§1.2)
3. Rich text sanitization (§1.3)
4. Bugs §2.1–2.6 (type switching, section deletion, duplicate keys, slider, blank wizard step, focus-opens-panel)

**Phase 2 — polish (this is where "clean and slick" is won)**
5. Theme tokens everywhere, dark mode (§3.1)
6. Inline form title + settings panel restructure (§3.2, §3.3)
7. Accessibility pass on labels/ARIA (§3.8)
8. Drag-to-reorder questions and sections (§3.4)
9. Pre-flight validation before Collect responses (§3.6)
10. `field.width` UI, or remove the model field (§3.8)

**Phase 3 — the dashboard**
11. Volume-over-time chart + KPI strip (§5.1) — cheapest big win
12. Chart-type switcher and real numeric distributions (§5.1)
13. Table view: column picker, sorting, virtualization (§5.1)
14. Cross-tab / segment-by (§5.1) — the differentiating feature
15. Drill-through from chart to filtered table

**Phase 4 — fields**
16. Fold Currency/Percent into Number; fold DateTime into Date; delete or build out Location (§4.1)
17. File upload (§4.2)
18. Likert matrix, Signature, Static content block (§4.2)
19. Lookup from another list (§4.2) — the strongest differentiator vs. Microsoft Forms
20. Multi-condition branching with AND/OR and numeric operators (§4.5)

**Phase 5 — scale & hygiene**
21. Paging + indexed `SFStatus` (§1.4)
22. Batched provisioning (§6)
23. O(n²) and clone-per-keystroke fixes (§6)
24. Unit tests for the pure utils (§7)
25. Delete dead code: draft scaffolding, unused props, unused SCSS (§2.15, §2.16)

**Also worth considering (not in the list above)**
- Form templates gallery — survey, event registration, feedback, IT request, expense claim, H&S check.
  Directly serves "options for any occasion" and makes the first-run experience feel finished.
- Save-and-resume drafts — the `SFStatus` column already exists for exactly this.
- Response limits, open/close dates, and one-response-per-person.
- QR code in the share dialog.
