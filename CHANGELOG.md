# Changelog

All notable changes to Smart Forms are documented here.

---

## [1.1.0] — 2026-09-30

### Added
- Localization: the interface is translated into 30 languages (the same set as Smart Org Chart) and follows the SharePoint page language
- Edit my response: with "Allow respondents to edit their response" on, respondents can reopen and update their latest submission
- Approval workflow: approve, reject or reset responses (with a comment and optional email to the respondent), an approval status column and filter, bulk actions, a dashboard tile, and approval columns in the CSV export
- Version history and restore: owners can list earlier saved versions of a form (date, author), preview the question count, and restore one with confirmation and one-click undo
- Right-to-left layout for Arabic, Hebrew, Persian and Urdu pages, and page-culture date formatting throughout
- Typed-name option on signature questions for keyboard users
- Drag and drop can place a question at the end of any section and into empty sections
- Designer question types, categories, presets, templates, operators and validation messages are localized

### Changed
- The saved-definition list now blocks inherited permissions and gives site members read-only access; only owners create or migrate it
- Response queries use `SFStatus eq 'Complete'` (older items are backfilled) so lists over 5,000 items keep working; unknown counts no longer close the form or skip the one-response rule
- Person questions and long column lists are read in extra requests, so forms with many Person questions load
- Copy link falls back to a textarea copy and then to a select-and-copy-manually prompt; in Microsoft Teams, share links use the SharePoint page URL
- Changing a question's type, or deleting or renaming choice options, removes branching rules that no longer fit and shows how many were removed
- Deleting a published question reserves its column name so a new question never adopts the old column
- Yes/No columns have no default, so skipped questions stay blank instead of counting as "No"
- Search on the Responses tab is debounced and bulk delete is faster
- Release workflow fails when the tag version does not match `solution.version`
- Package description says 24 question types; documentation corrected (share links require signed-in users with at least read access; URL prefill is not a security boundary)

### Fixed
- Resuming a draft no longer loses attachments or signatures, and answers cleared since the draft was saved no longer survive the final submit; the form waits for the draft to load and guards against double submits
- Failed attachment uploads are reported to the respondent, and duplicate file names are detected case-insensitively
- Date and time branching rules and URL prefill use local dates and HH:mm times; prefilled choice values must match an option
- Hidden questions no longer drive branching or calculated fields; formulas accept very large and very small numbers
- Clicking a chart bar (rating stars, "Other" slice, multi-select answers) now filters to exactly the answers counted; date and time charts are chronological
- CSV export keeps negative numbers numeric and guards more formula-injection prefixes
- HTML sanitizer now blocks lowercase SVG/MathML elements and protocol-relative links
- Loading the saved definition no longer overwrites edits made before the load returns

---

## [1.0.0] — 2026-09-28

### Added
- Initial release.
