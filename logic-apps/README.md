# Logic App workflows

The five Azure Logic App (Standard) workflows the web parts depend on. Each folder holds one workflow's `workflow.json`, in the same layout a Logic App Standard project uses. The folder name is the workflow name.

Four are HTTP-triggered "elevated write" endpoints, called through Azure API Management (APIM) by the web parts. They write to SharePoint with the Logic App's own connection, because staff only hold Read (or Add Items) on report items. The fifth keeps each report's item-level permissions in sync.

| Workflow | Trigger | Called from | Who it allows | Connection |
| --- | --- | --- | --- | --- |
| `Data_Breach_Send_Email` | HTTP `When_an_HTTP_request_is_received` | `sendEmailFromMailbox` in `DataBreachForm.tsx` | Anyone with the APIM key (no identity check) | Office 365 Outlook |
| `Data_Breach_Append_Comm_Log` | HTTP `When_a_HTTP_request_is_received` | `appendCommLogEntryElevated` in `SharePointService.ts` | The report's author or person responsible, or a member of the Reviewers/Owners groups | SharePoint |
| `Data_Breach_Update_Panel_Recommendation` | HTTP `When_a_HTTP_request_is_received` | `updatePanelRecommendationsElevated` in `SharePointService.ts` | Members of the Reviewers/Owners groups | SharePoint |
| `Data_Breach_Upload_Attachment` | HTTP `When_a_HTTP_request_is_received` | `uploadAttachmentElevated` in `SharePointService.ts` | The report's author | SharePoint |
| `Data_Breach_Permission_Handler` | SharePoint item created or modified (polls every minute) | — | — | SharePoint |

Note the two trigger names differ (`an` vs `a`). The trigger name is part of each workflow's callback URL, so your APIM policies must use the exact name.

## Values to replace

Search each `workflow.json` for `REPLACE_WITH_` and fill in your own values before deploying:

| Placeholder | Used in | Value |
| --- | --- | --- |
| `REPLACE_WITH_SHAREPOINT_SITE_URL` | All except Send Email | Full URL of the site holding the list, e.g. `https://yourtenant.sharepoint.com/sites/DataBreach/` |
| `REPLACE_WITH_LIST_ID` | Upload Attachment, Permission Sync | The list's GUID (List settings → the `List=` value in the URL) |
| `REPLACE_WITH_FROM_MAILBOX` | Send Email | The shared mailbox emails are sent from. The connection's account needs **Send As** permission on it |

These names are written into the workflows as the defaults the web parts also use. Change them in both places if yours differ:

- **List:** `Data Breaches`. The SharePoint entity type `SP.Data.Data_x0020_BreachesListItem` used in the update actions is derived from the list name, so it changes too.
- **Groups:** `Data Breach Owners` (Contribute on every report) and `Data Breach Reviewers` (Read on every report). These must match `SP_ADMIN_GROUPS` and `SP_REVIEWER_GROUPS` in `.env`.
- **Permission levels:** the built-in `Read` and `Contribute`.

## Connections

The workflows refer to their connections by name: `sharepointonline` and `office365`. When you create the workflows, either create connections with those names, or open each action in the designer and select your own connection, which updates the name for you.

- **SharePoint:** an account with **Full Control** on the site, since the workflows break permission inheritance and grant access on report items.
- **Office 365 Outlook:** an account with **Send As** on the `From` mailbox.

## Deploying

1. Create a Logic App (Standard), or use an existing one.
2. For each folder here, create a **stateful** workflow with **exactly the folder's name**, open **Code view**, and paste in `workflow.json` with your values filled in.
3. Set up the connections as above and save.
4. Copy each HTTP workflow's callback URL from its trigger. In APIM, create one operation per workflow that rewrites to `/api/<WorkflowName>/triggers/<TriggerName>/invoke` and sets the `api-version`, `sp`, `sv` and `sig` query parameters from that URL. Store `sig` in a secret Named Value, and protect each operation with its own key (`check-header` against a Named Value).
5. Enter each operation's APIM URL and key in the web parts' property panes.

Read the **Security model** section of the main README before deploying. These endpoints are designed to run on an internal network only.

## Optional changes

- **Sensitivity label:** the Send Email action can apply a Microsoft Purview sensitivity label by adding `"Sensitivity": "<label GUID>"` to its body.
- **Restrict recipients:** Send Email will send to any address it is given. Consider adding a check that every recipient belongs to your own domain(s) before the send action.
- **Retries:** Send Email retries a failed send 5 times at 1-minute intervals. Callers stop waiting long before that, so a shorter policy (e.g. 3 × 10 seconds) gives clearer failures.
