# SharePoint Smart Forms — SPFx Form Builder Web Part

[![Website](https://img.shields.io/badge/Website-sharepointsmartsolutions.com-blue)](https://sharepointsmartsolutions.com/smart-forms) [![Download](https://img.shields.io/badge/Download-Latest%20Release-CA5010?logo=github&logoColor=white)](../../releases/latest) [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Smart Forms brings a Microsoft Forms-style experience to SharePoint — form owners visually design forms with 24 question types (including several Forms doesn't have, like person pickers, live lookups from another list, calculated totals, signatures, and Likert grids), share a fill-in link with one click, and analyse responses on a built-in dashboard. Every submission lands in a regular SharePoint/Microsoft Lists list with real columns, so your data is never locked inside the form tool.

![SPFx](https://img.shields.io/badge/SPFx-1.20.0-0078D4?logo=microsoft&logoColor=white) ![React](https://img.shields.io/badge/React-17-61DAFB?logo=react&logoColor=black) ![TypeScript](https://img.shields.io/badge/TypeScript-4.7-3178C6?logo=typescript&logoColor=white) ![PnPjs](https://img.shields.io/badge/PnPjs-3.26-217346) ![Fluent UI](https://img.shields.io/badge/Fluent%20UI-8-742774)

---

## Features

### Form designer

| Feature | Description |
|---|---|
| **Edit-in-place designer** | The editor *is* the form — click any question card and edit it right there: type into the label, edit options inline, flip Required in the card footer. The form title and description sit at the top of the canvas, not behind a panel |
| **Drag to reorder** | Drag questions by their grip, or from the outline rail — including across section boundaries. Arrow keys work on the grip for keyboard reordering |
| **Insert anywhere** | Every question's **…** menu has an **Insert question below** picker, so a new question can be dropped in right after any existing one instead of adding at the end and dragging it into place |
| **Outline rail** | A live table of contents for long forms: jump to any question, search by text, and see at a glance which questions have branching or a problem |
| **24 question types** | Short/long/formatted text · number (plain, currency or percentage) · calculated · date (with optional time) · time · choice · image choice · **lookup from another list** · yes/no · consent · ranking · rating (stars/hearts/thumbs, optional halves) · opinion scale · slider · **Likert grid** · email · phone · address · link · person · **file upload** · **signature** · text block |
| **Presets** | One-click NPS (0–10), satisfaction (1–5), currency and percentage questions built on the general types |
| **Branching with AND/OR** | Show a question only when a group of conditions is met — `is`, `is not`, `contains`, `is answered`, plus numeric `greater than` / `between` and date `before` / `after`. The rule is restated in plain English as you build it |
| **Validation** | Required fields with custom messages, email/phone/URL checks, numeric and scale bounds, min/max selections, regex patterns, file type and size limits — all enforced before submit |
| **Pre-flight check** | **Collect responses** first checks for empty option lists, broken formulas, branching pointing at deleted questions and unnamed questions, with a "Go to" jump link for each |
| **Templates** | Eight starting points — customer feedback, event registration, IT request, expense claim, employee pulse, safety inspection, room booking, or blank |
| **Automatic column provisioning** | Creates the matching SharePoint columns behind the scenes (clean internal names, grouped under "Smart Forms"), several at a time rather than one-by-one |

### Form filler experience

| Feature | Description |
|---|---|
| **Collect responses** | One click provisions the list columns and opens a share dialog with a copy-able link (`?sfview=fill`) that shows just the form — plus a reminder of the list permission respondents need |
| **Non-destructive preview** | Preview validates and shows the confirmation screen without writing anything to the response list |
| **URL prefill** | Add `&<ColumnName>=value` to a share link to pre-answer questions (e.g. `&SFDepartment=Finance`); questions can also be marked read-only so a prefilled value can't be changed |
| **Theme-aware** | Every colour derives from the SharePoint site theme, so the form is legible on light and dark sites. The owner's accent colour is contrast-corrected automatically |
| **Accessible by construction** | Each question is a labelled group with proper `aria-describedby`, error and required semantics; scales, ratings and Likert grids are real radio/slider patterns with arrow-key support; an error summary links straight to the questions that need attention |
| **Wizard mode** | One section per step with clickable step markers, a progress bar, per-step validation, and no more blank steps when branching hides a whole section |
| **Save and resume** | Optionally let respondents save a partial response and finish it later |
| **Access controls** | Open/close dates, a response cap, and one-response-per-person |
| **Smart choice display** | Choice questions render inline when the option list is short and switch to a dropdown for long lists; options can shuffle per respondent, and a write-in "Other" is one toggle |
| **Files and signatures** | Drag-and-drop uploads with type/size limits, and a pen/touch signature pad — both stored as attachments on the response item |
| **Calculated totals** | A read-only question that works out a total from other answers (`{Quantity} * {Unit price}`), updating as you type |

### Results (owners only)

| Feature | Description |
|---|---|
| **Dashboard** | Headline KPIs (total, last 7 days with a week-on-week delta, today, unique respondents, completion rate, median time to complete), a response-volume timeline by day/week/month, and auto-selected highlight tiles for the questions worth watching |
| **Configurable tiles** | Pin, hide, reorder and re-chart any question — bar, column, donut, gauge, histogram, statistics, word frequency or table. The layout is saved with the form |
| **Compare by** | Split every summary card by any choice, yes/no, person or rating question — "NPS by department", "satisfaction by region" — without leaving the page |
| **Drill-through** | Click a bar or donut segment to filter the response table to exactly those respondents |
| **One filter bar** | Search, date range and answer filters apply to the dashboard, summary and table together, so all three always report the same number |
| **Response table** | Every question available as a column with a picker and custom order, sortable columns, virtualised rows, multi-select with bulk delete |
| **Individual responses** | Prev/next through the filtered set, read either as a labelled list or as the form itself, with attachment links, signature previews, and a print stylesheet |
| **Charts built for reading** | Palettes are validated for colourblind separation and contrast against the actual surface in both light and dark mode; every chart carries visible value labels and a table view |
| **Email notifications** | Every response emailed to the recipients you choose — full answers, formatted, with a **View in Microsoft Lists** link. No Power Automate, no connectors, no premium licensing |
| **Respondent receipts** | Optionally email each respondent a copy of their own answers |
| **CSV export** | One-click export of the (filtered) responses, Excel-ready with UTF-8 BOM and formula-injection guarding |
| **Permission aware** | Visitors only ever see the form — the designer and results appear only for users with Manage Lists permission |

---

## Prerequisites (for Development Only)

| Requirement | Detail |
|---|---|
| **Node.js** | 18.x (`>=18.17.1 <19.0.0`) |
| **SharePoint** | Online (Microsoft 365) |
| **SPFx** | 1.20.0 |
| **Permissions to deploy** | Site Owner or above |

---

## Development Setup

```bash
# install dependencies
npm install

# point the workbench at your tenant
cp config/serve.json.example config/serve.json
# then edit config/serve.json and replace <your-tenant> with your tenant name

# trust the dev certificate (first time only)
npm run trust-dev-cert

# start the local workbench
npm run serve
```

### Tests

```bash
npm test
```

Compiles the pure utility modules to CommonJS and runs them on Node — no extra
dependencies, no browser. Covers validation, branching, the formula evaluator,
schema migration, value round-tripping, CSV, analytics and the chart colour
rules.

---

## Build & Deploy

```bash
# production build and package
npm run ship

# output
# sharepoint/solution/sharepoint-smart-forms.sppkg
```

1. Upload `sharepoint/solution/sharepoint-smart-forms.sppkg` to your tenant **App Catalog** (or a site collection app catalog)
2. When prompted, choose whether to make the web part available to all sites (`skipFeatureDeployment` is enabled, so no site-level install is needed)
3. Edit a page, add the **Smart Forms** web part from the **Other** group
4. In the property pane, pick an existing list for responses — or type a name and click **Create list**
5. Build your form on the page, or pick a template from the lightbulb button
6. **Save or publish the page** — the form definition is stored with the page
7. Click **Collect responses** and send the copied link to anyone who should fill in the form

> **The form lives with the page.** SPFx keeps web part properties in memory until
> the page is saved, so a form built and then abandoned without publishing is lost.
> The web part shows a warning banner and a save-state chip while changes are
> unsaved — publish the page to keep your work.

---

## Configuration

Almost everything is edited in place on the page. The property pane only picks the response list.

### Response list

| Setting | Description |
|---|---|
| **Save responses to this list** | Dropdown of the site's visible generic lists; each submission becomes a list item |
| **Or create a new list** | Type a name and click **Create list** to create and select a fresh list without leaving the page |

### Form settings (the gear button on the page)

| Tab | Contains |
|---|---|
| **Basics** | Title, description, single-page or wizard layout, question numbering, question shuffling |
| **Appearance** | Accent colour (eight presets plus a custom picker) and header icon |
| **After submit** | Submit button text, thank-you title and message, "Submit another response" |
| **Notifications** | Who gets emailed each response, and respondent receipts |
| **Access** | Open/close dates, response cap, one-response-per-person, save-and-resume, closed message |

Per-question settings — per-type options, branching, and validation (placeholders, defaults, required messages, patterns) — live in the question's **…** menu under **Branching, validation & more**, split into **Options** / **Branching** / **Validation** tabs.

Emails are sent through SharePoint's built-in send-email API (`no-reply@sharepointonline.com`), so recipients must be **users in your Microsoft 365 organization** — external addresses are dropped. Notification failures never block a submission.

### Sharing

| Link | Behavior |
|---|---|
| Page URL | Visitors see just the form; owners see the designer and results |
| Page URL + `?sfview=fill` (what **Collect responses** copies) | Everyone — including owners — sees just the form, ready to fill in |
| Share link + `&<ColumnName>=value` | Pre-answers a question; use the question's internal column name, e.g. `&SFDepartment=Finance` |

Respondents need permission to **add items** to the response list. If someone reports an error on submit, that is almost always the cause.

---

## Project Structure

```
SharePointSmartForms/
├── config/
│   ├── package-solution.json        # solution id, version, .sppkg path
│   ├── serve.json.example           # workbench URL template (copy to serve.json)
│   ├── tsconfig.test.json           # CommonJS re-compile for the test runner
│   └── config.json                  # bundle definition
├── tests/
│   ├── harness.ts                   # dependency-free assertions and runner
│   └── index.ts                     # unit tests for the pure modules
├── src/webparts/smartForms/
│   ├── SmartFormsWebPart.ts         # web part entry, theme wiring, list picker/creator
│   ├── SmartFormsWebPart.manifest.json
│   ├── models/
│   │   ├── index.ts                 # schema, field types, dashboard settings, v1→v2 migration
│   │   └── templates.ts             # eight starter forms
│   ├── hooks/
│   │   └── useDragList.ts           # shared drag-to-reorder behaviour
│   ├── services/
│   │   └── SharePointService.ts     # PnPjs layer: provisioning, paging, attachments, lookups, drafts
│   ├── utils/
│   │   ├── formUtils.ts             # validation, condition engine, formatting, CSV, pre-flight
│   │   ├── analytics.ts             # distributions, NPS, Likert, histograms, timeline, KPIs, segments
│   │   ├── formula.ts               # shunting-yard evaluator for calculated fields (no eval)
│   │   ├── sanitizeHtml.ts          # allowlist sanitizer for stored rich text
│   │   ├── theme.ts                 # site theme → CSS custom properties, validated chart palettes
│   │   └── spFieldXml.ts            # form field → SharePoint column schema XML
│   └── components/
│       ├── _tokens.scss             # shared design tokens and mixins
│       ├── SmartForms.tsx           # root: tabs, save state, pre-flight, share, templates
│       ├── charts/                  # inline SVG bar/column/donut/area/gauge/diverging/sparkline
│       ├── form/                    # FormRenderer + one component per question type
│       ├── designer/                # canvas, outline rail, field settings, form settings
│       └── responses/               # dashboard, summary, table, detail panel, QuestionChart
├── gulpfile.js
└── package.json
```

---

## Key Dependencies

| Package | Purpose |
|---|---|
| `@fluentui/react` (v8) | UI controls — pickers, panels, grid, command bars |
| `@pnp/sp` + `@pnp/core` (v3.26) | SharePoint REST data layer — lists, fields, items, attachments, people search |
| `@microsoft/sp-component-base` | Site theme provider |
| `@microsoft/sp-webpart-base` | SPFx web part framework |
| `react` / `react-dom` (17) | Component rendering |

Charts are hand-built inline SVG rather than a charting library, to keep the SPFx bundle small.

---

## Troubleshooting

- **My form disappeared** — the page was never saved after editing. SPFx stores the form with the page; watch for the amber "publish the page" banner
- **Submissions fail for colleagues but work for me** — respondents need **Add Items** permission on the response list
- **"Responses could not be loaded"** — the list columns don't exist yet; click **Collect responses** once so they are created
- **The designer and results tabs are missing** — they only appear for users with **Manage Lists** permission on the response list
- **"The list columns could not be created"** — the signed-in user needs Manage Lists on the response list
- **"already has a list column of a different type"** — a question's type was changed after publishing. SharePoint can't retype a column in place; rename the question so a fresh column is created
- **"Change type" is disabled on a question** — the backing column already exists; duplicate the question to get a different type
- **A percentage shows as a plain number in the list view** — intentional. The column stores the number the respondent typed (45, not 0.45) so exports, Power BI and the list view all agree; the form adds the `%` on display

### Getting debug output

The web part logs to the browser console under the `[SmartForms]` prefix:

- **Errors always log** — a failed list load, publish, submit or draft save prints the real exception under `[SmartForms]`, even though the on-screen message stays generic. Filter the console to `smartforms` to see just these.
- **Verbose tracing is opt-in.** Run `localStorage.setItem('smartFormsDebug', '1')` in the console and reload to trace list loads, provisioning, and non-fatal warnings (a column that already existed under a different name, a view that couldn't be updated, and similar best-effort steps that don't stop the form working). Run `localStorage.removeItem('smartFormsDebug')` to turn it back off.
- **A render crash shows a red error panel** with the message and stack instead of leaving the web part blank, so a failure is always visible somewhere — even if nothing was logged in time to see it.

If the web part shows nothing at all and the console has no `[SmartForms]` lines, the bundle likely never loaded — check that `gulp serve` is running and that the workbench URL includes `debugManifestsFile=https://localhost:4321/temp/manifests.js`.

---

## Limitations

- Notification emails reach organization users only — SharePoint's send-email API cannot deliver to external addresses (use a Power Automate flow on the response list if you need that)
- Results views load the newest 20,000 responses; beyond that the views show a notice and the list itself remains complete
- A question's type (and a choice question's single/multi setting) is locked once its column has been published
- Lookup questions store the resolved option *text*, not a link to the source item — responses stay readable if the source list changes, but a later rename is not reflected in existing responses
- Addresses are stored as one formatted line rather than separate columns, so address parts can't be analysed independently
- Access rules (open/close dates, response caps, one-per-person) are enforced by the form, not by SharePoint — anyone with Add Items permission could still add an item directly through the list
- The rich text editor covers bold, italic, underline and lists — not tables or images
- Save-and-resume is available on single-page forms, not wizard forms
- Deleting a question removes it from the form but leaves the SharePoint column (and its data) in the list

---

## License

[MIT](LICENSE) © 2026 Sean Regan
