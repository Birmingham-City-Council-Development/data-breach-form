# Data Breach Form

SPFx 1.20.0 web part for reporting data breaches. Built with React 17 and GOV.UK Frontend styling. Submits form data directly to a SharePoint list.

## Architecture

This is an **SPFx (SharePoint Framework) 1.20.0** project. If you're used to standard React apps (Create React App, Vite, Next.js), the key differences are:

**Build system** — Instead of Webpack or Vite, the build is driven by **Gulp** via Microsoft's SPFx build rig (`@microsoft/sp-build-web`). The main commands are `gulp bundle` and `gulp package-solution` rather than `npm run build`.

**Component model** — There is no single `index.tsx` root. Each webpart is a class extending `BaseClientSideWebPart` with its own `.manifest.json` that registers it with SharePoint. This project has two: `DataBreachFormWebPart` (the submission form) and `ReportViewerWebPart` (the admin report viewer).

**Deployment** — The project doesn't deploy to a web server or CDN. It builds a `.sppkg` (SharePoint Package) file that is uploaded to a SharePoint App Catalog and installed on a site.

**SharePoint data access** — There is no custom API layer. The project uses **PnPjs** (`@pnp/sp`) to read and write SharePoint list items directly, and the `@microsoft/sp-*` framework packages provide context such as the current user, site URL, and property pane configuration.

**React version** — React is pinned at **17.0.1** because SPFx 1.20.0 requires it. It cannot be upgraded independently of the SPFx version.

**Dev environment** — Instead of `localhost:3000`, development runs against a live SharePoint workbench URL (`/_layouts/workbench.aspx`), so a connection to your SharePoint tenant is required during development.

**Styling** — SCSS modules (SPFx standard) and **GOV.UK Frontend** (`govuk-frontend`) are layered together. GOV.UK Frontend provides the UK government design system components and class conventions.

## Prerequisites

- Node.js 18.x
- A SharePoint site to host the solution (e.g. `https://yourtenant.sharepoint.com/sites/DataBreach`)
- A `Data Breaches` list on that site

## Configuration

Deployment settings live in a `.env` file at the project root, which is gitignored. Copy the template and fill it in:

```
cp .env.example .env
```

| Setting | Purpose |
| ---- | ---- |
| `ORG_NAME` | Organisation name in email headers/footers (blank to omit) |
| `TEAM_NAME` | Team handling reports, used in "on behalf of" attributions in the communication log and PDF |
| `REVIEW_TEAM_NAME` | Team named in the pre-filled "Reason for decision" wording (defaults to `TEAM_NAME`) |
| `ADMIN_NOTIFICATION_EMAIL` | Central team mailbox: new-report and case-closed notifications, fallback reply address |
| `SP_SITE_URL` | Site URL, used to open the workbench on `gulp serve` |
| `SP_LIST_NAME` | Display name of the SharePoint list |
| `SP_ADMIN_GROUPS` / `SP_REVIEWER_GROUPS` | Comma-separated SharePoint group names for the Owner and Reviewer roles |
| `LIVE_SERVICE_URL` | Production form page, linked from the `TEST_MODE` banner |
| `MAX_ATTACHMENT_MB` | Per-file upload limit on the form (default 20) |
| `ALLOWED_ATTACHMENT_EXTENSIONS` | Comma-separated extensions the form accepts (blank for the default list) |
| `SENDMAIL_URL`, `APPEND_COMM_LOG_URL`, `UPDATE_PANEL_RECOMMENDATIONS_URL`, `UPLOAD_ATTACHMENT_URL` | Optional default APIM URLs; a URL set in the web part property pane takes priority |

`gulpfile.js` reads `.env` and bakes the values into the bundle at build time (exposed to the code through `src/webparts/dataBreachForm/config.ts`), so **rebuild after changing it**. Real environment variables override `.env`, which is useful in CI.

**API keys don't go in `.env`**: they would end up inside the `.sppkg`. Enter them in each web part's property pane instead.

Dropdown options, including your organisation's directorates and divisions, are in `src/webparts/dataBreachForm/components/formOptions.ts`.

## Local development

Install dependencies:
```
npm install
```

Start the local dev server (opens the SharePoint workbench):
```
npx gulp serve
```

The workbench will open at `<SP_SITE_URL>/_layouts/workbench.aspx`.

Add the **Data Breach Form** web part from the toolbox to test it. Submissions will write to the real `Data Breaches` list.

## Packaging for deployment

Run the following commands in order:

```
npx gulp bundle --ship
npx gulp package-solution --ship
```

The deployable file will be created at:
```
sharepoint/solution/data-breach-form.sppkg
```

## Deployment

1. Go to your SharePoint site
2. Go to **Site contents** and open the **Apps for SharePoint** library
3. Upload `data-breach-form.sppkg` and click **Deploy** when prompted
4. Go to **Site contents → Add an app** and install **data-breach-form**
5. Add the web part to a page via **Pages → Edit → + → Data Breach Form**

> If the **Apps for SharePoint** library does not exist, a SharePoint admin needs to enable the site collection app catalog. In PowerShell 7:
> ```powershell
> Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
> Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/DataBreach" -Interactive
> Add-PnPSiteCollectionAppCatalog -Site "https://yourtenant.sharepoint.com/sites/DataBreach"
> ```

## Email sending

There are **two separate, unrelated code paths** that send email — don't assume one describes the other's behavior:

1. **Standalone `ReportViewerWebPart`** (when placed on a page by itself) sends via **Microsoft Graph** directly, using the mailbox configured in the web part's "Shared mailbox email address" property (e.g. `dataprotection@example.gov.uk`). That mailbox must grant the web part's service account Send As permission.

2. **The embedded report viewer inside the Data Breach Form web part's "View reports" mode** sends through the `sendmail` Logic App operation (see "Elevated write endpoints" below). The web part only posts `{ to, subject, body }` — the actual "from" mailbox is configured server-side in the Logic App itself, not by anything in this codebase. The "Sender email address" property pane field on the Data Breach Form web part does **not** control this — it's only used as display text (who a reply appears to come from in the UI), not a functional sender override.

`TEST_MODE` in `DataBreachForm.tsx` prefixes `[TEST]` on subjects, adds a banner to bodies for anything sent via the Logic App path, and shows an on-screen warning linking to `LIVE_SERVICE_URL`. Set it to `true` for a test deployment and `false` for production.

## Roles and permissions

Three distinct tiers, each backed by a different SharePoint group and a different level of item access:

| Role | SharePoint group | Item-level access | Capability |
| ---- | ---- | ---- | ---- |
| **Owner** | `Data Breach Owners` | Contribute (incl. Delete) | Full admin: edit every field, delete a report, side nav, everything |
| **Reviewer** | `Data Breach Reviewers` | Read only | View every report, edit Panel recommendations, change Status, Request more information — everything via elevated Logic App endpoints, never direct edit rights |
| **Named individual** (submitter / reporting officer / person responsible / line manager) | — resolved per-item from the report's own fields | Read only | View their own report(s); only the **person responsible** gets the reply action |

Owners are the technical/developer role, not business reviewers — that split is deliberate. Reviewers and named individuals never hold raw Edit rights on a list item; every write they can make goes through an elevated Logic App endpoint that verifies the caller server-side before writing, since SharePoint has no way to grant "edit rights on just one field."

## Elevated write endpoints (APIM + Logic Apps)

Four Logic App workflows, all hosted in one Logic App Standard resource, fronted by one shared APIM API (each operation independently secured with its own subscription key — not APIM's native per-Product mechanism, since that would apply to every operation at once):

- **`sendmail`** — sends an email from the shared mailbox configured in the Logic App itself
- **`append-comm-log`** — appends one entry to the report's communication log (conflict-safe, with retry); callable by the item's Author/person responsible, or any Reviewer/Owner
- **`update-panel-recommendations`** — updates the Panel recommendations fields and Status; callable by any Reviewer or Owner
- **`upload-attachment`** — attaches a file to a report on behalf of its submitter, who holds Add Items but not Edit on the list

Each operation's APIM URL and key are entered in the web part's property pane. The workflows are in [`logic-apps/`](logic-apps/README.md), with setup instructions. Each one checks the caller's identity or SharePoint group membership before writing with its own elevated SharePoint connection.

## Security model

Read this before deploying. The elevated endpoints are designed to sit on an **internal network**, and the current design relies on that:

- **API keys are visible in the browser.** The web part has to send them from the client, so anyone who can open the page can read them with developer tools. They are a shared secret, not a strong control.
- **Caller identity comes from the request body.** Each request includes a `callerEmail`, which the Logic Apps use for their Author/person-responsible/group checks. Someone holding a key could send another person's email address.

So **only deploy APIM and the Logic Apps on an internal network** (APIM in internal VNet mode, Logic Apps not publicly reachable), so that only people already on your network can call them, and keep logs of who calls them.

To remove both limitations, validate Microsoft Entra ID tokens instead: call the endpoints from SPFx with `AadHttpClient`, add APIM's `validate-jwt` policy, and pass the caller's email from the validated token's claims to the Logic Apps in place of `callerEmail`. This does not restrict who can use the form; it proves who is making each call.

Other recommended APIM settings: a per-IP rate limit, a request-size cap above the attachment limit, and storing each Logic App trigger's `sig` in a secret Named Value rather than in the policy text.

## SharePoint list setup

The web part writes to the `Data Breaches` list using the column internal names defined in `src/webparts/dataBreachForm/services/IBreachListItem.ts`.

The report reference (e.g. `DBN-55`) is computed client-side from the list item's own `Id`, so no ID column or flow is needed.

Base list/site permission should be kept minimal for everyone — all real per-item access (who can see/edit which specific report) comes from the item-level grants described in "Roles and permissions" above, not from a broad base permission level. Regular staff (submitters, reviewers included) only need a custom permission level with **Add Items + View Items** — no Edit Items, no Manage Lists, no Delete. Even the Owners (admin/dev) group should be kept off Full Control/Site Owner at the base level; their real capability comes entirely from the per-item Contribute grant `syncItemAccess` applies, not from broad site power.

## Solution

| File | Purpose |
| ---- | ------- |
| `.env.example` | Template for deployment settings (copy to `.env`) |
| `src/webparts/dataBreachForm/config.ts` | Exposes the `.env` settings to the code |
| `src/webparts/dataBreachForm/components/DataBreachForm.tsx` | Main form component |
| `src/webparts/dataBreachForm/components/IFormState.ts` | Form state interface |
| `src/webparts/dataBreachForm/components/formOptions.ts` | All dropdown options and directorates |
| `src/webparts/dataBreachForm/services/SharePointService.ts` | PnPjs SharePoint write service |
| `src/webparts/dataBreachForm/services/IBreachListItem.ts` | SharePoint column internal name mapping |

## Licence

Released under the [MIT licence](LICENSE). Copyright (c) 2026 Birmingham City Council.
