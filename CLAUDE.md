# OptikTV / Entertainment On Green Dashboard

## What this project is
A Google Apps Script web app (`doGet()` serves `index.html` via `HtmlService`). All UI is a single `index.html` file. Server-side logic lives in `.gs` files deployed as a GAS project.

---

## File map

| File | Purpose |
|------|---------|
| `index.html` | Entire front-end (~8000 lines): tabs, tables, filters, charts, modals |
| `Core.gs` | Config constants only — channels, keywords, feature flags, filter IDs |
| `Code_util.gs` | GAS entry point (`doGet`), utility functions, `verifyPin()` |
| `Code_slack.gs` | Slack integration: fetch executive status, send messages, format text |
| `Code_defects.gs` | Jira defects fetch + normalization |
| `Code_epic.gs` | Jira epics/work-items fetch + normalization |
| `Code_cost.gs` | Jira assessments/cost fetch + normalization |
| `Code_timeline.gs` | Jira timeline fetch |
| `GREEN_code.gs` | Legacy/misc Jira helpers |

---

## Architecture notes

- **No build step** — edit files directly, deploy via Apps Script editor or `clasp push`
- **Script Properties** (set in Apps Script project settings) store secrets:
  - `SLACK_BOT_TOKEN` — Slack bot token (needs `channels:history`, `users:read` scopes)
  - `JIRA_SERVICE_ACCOUNT_ID`, `JIRA_SECRET_KEY`, `JIRA_CLOUD_ID`
  - `GCHAT_WEBHOOK_TESTSPACE`, `GCHAT_WEBHOOK_CIO`, `GCHAT_WEBHOOK_BUSINESS`
  - `PIN` — PIN required before sending to production Slack/GChat destinations
- **localStorage** (`vogAssessmentsJiraData`) persists loaded Jira data client-side across refreshes
- **`google.script.run`** is the client→server bridge (GAS-specific, not fetch/XHR)

---

## Key decisions & patterns

### Config changes → only touch Core.gs
- Slack channels: `CONFIG.SLACK.channels`
- Executive status search keyword: `CONFIG.SLACK.executiveStatusKeyword` (currently `'entertainment 5.0 program'`, case-insensitive match against last 50 messages in `#prod-green-commerce`)
- Google Chat spaces: `CONFIG.GCHAT.spaces`
- Feature flags (enable/disable tabs): `CONFIG.FEATURE_FLAGS`

### PIN protection
- Any send to a non-`isTest` Slack channel or GChat space requires PIN entry
- `verifyPin(pin)` in `Code_util.gs` — returns `{success: bool}`, never echoes the PIN
- Test destinations marked with `isTest: true` in `Core.gs`

### Tab structure in index.html
Each tab has its own: data array, filtered array, sort config, filter state, and `update*Table()` function.

| Tab | Data var | Filter fn | Sort config |
|-----|----------|-----------|-------------|
| Work Items | `epicsData` / `epicsFilteredData` | `applyEpicsFilters()` | `epicsSortConfig` |
| Defects | `defectsData` / `filteredData` | `applyFilters()` | (inline) |
| Assessment | `assessmentsData` / `assessmentsFilteredData` | `applyAssessmentsFilters()` | `assessmentsSortConfig` |

### Defects tab metrics (two-panel layout)
- **Left — Defect Breakdown** (`#defects-breakdown-panel`): clickable severity/status rows; clicking filters the table to that category. State: `defectBreakdownFilter`.
- **Right — Quick Filters** (`#defects-quick-filters-panel`): 6 clickable cards (Age>10, Unassigned, Open, Overdue, No Due Date, Due Soon). State: `defectQuickFilter`.
- Both filters are AND-combined with dropdown filters in `applyFilters()`.
- Only rows with count > 0 are rendered (matches Work Item Breakdown behaviour).

### Summary column tooltip
- Shared helpers `showSummaryTooltip(text)` / `hideSummaryTooltip()` defined at top of `<script>` block
- `#summary-tooltip` div lives just before `</body>` (must be lazy-fetched — the div doesn't exist when the script tag is first parsed)
- Wired up in Work Items, Assessments, and Defects `update*Table()` functions via `mouseenter`/`mouseleave`

### Sorting
- Every sortable column needs three things: (1) click listener registered, (2) entry in `update*SortIndicators()`, (3) `case` in `apply*Sorting()` switch
- Common bug: adding a new column header but forgetting one of the three places

### Slack executive status fetch flow
1. `generateExecutiveSummary()` (index.html) calls `fetchExecutiveStatusFromSlack()` via `google.script.run`
2. Fetches last 50 messages from `C03UMGV7DDE` (prod-green-commerce)
3. Returns first message matching `CONFIG.SLACK.executiveStatusKeyword` (case-insensitive)
4. Fallback: most recent human message + note shown in yellow banner
5. `resolveSlackUserMentions()` → `formatSlackText()` → `renderSlackText()` pipeline
6. Warning banner (`#executive-fetch-note`) shown if any `@mention` UID fails to resolve

---

## Development branch
`claude/refactor-application-mATQz` (on `nicks0915/main` — old repo)

When working in the new repo (`telus/OptikTVDashboard`), create a new feature branch from main.

---

## Recent changes (for context)
- Executive status keyword made configurable in `Core.gs`
- Defects tab fully redesigned: Breakdown panel + Quick Filters panel
- Summary column hover tooltip added to all 3 tabs
- Sorting fixed: assessments labels column, work items labels column, missing `updateAssessmentsSortIndicators()` function
- Defect breakdown hides zero-count rows (matches Work Item Breakdown)
- PIN protection on all production Slack/GChat sends
- Slack `@mention` resolution diagnostics + yellow warning banner
