---
name: m365-graph
description: Search and read the organisation's Microsoft 365 content — Outlook mail, SharePoint and OneDrive documents, Teams chats, the user's calendar — through Microsoft Graph with the signed-in user's own delegated permissions. List the signed-in user's OneDrive with /me/drive/root/children, open a folder from that list, and read a Word, Excel or PDF file with /me/drive/items/<id>/content. Use when the user asks to find something that lives in the cloud rather than on this machine ("keresd meg a levelet", "mi van a OneDrive-omon", "mit ír a dokumentum"). Do NOT use for a document the user has OPEN in Word/Excel/PowerPoint — that is the office-live skill. Do NOT use for local files.
version: 1.0.0
author: Integritás Hatóság
platforms: [windows]
---

# m365-graph

To list the signed-in user's OneDrive, call `m365_get` with
`/me/drive/root/children`. The answer is `count`, `folders` and `files`.
Report those fields and do not count the items again. Open a folder from
that list with `/me/drive/items/<id>/children`. A folder addressed by its
name is `/me/drive/root:/Folder name:/children`.

To read a Word, Excel or PDF file, call `m365_get` with
`/me/drive/items/<id>/content` and summarize `text`. A `.docx` is read
locally (`text_source` is `docx`). A `.xlsx` is read locally (`xlsx`): the
sheet name, then rows. A legacy `.doc` or `.xls` is not extracted; ask for
the `.docx` or `.xlsx` copy.

A PDF page with a text layer is read locally by PyMuPDF (`text_source`
contains `text-layer`) and never leaves the machine. A scanned page goes to
the internal llama-swap OCR model (`local-vlm`). Only if that page is still
empty does the raw PDF go to Azure Document Intelligence (`azure-di`), and
only if that call fails does pdfminer read the file locally (`pdfminer`).
Pages that are still empty are sent as images to Azure OpenAI vision
(`azure-vision`). Do not open the SharePoint link in the browser, and do
not call `/preview`: the browser viewer does not expose the text layer, and
`/preview` returns 403.

Searches and reads Microsoft 365 through one file:
`${HERMES_SKILL_DIR}/scripts/microsoft_graph.py`. The user signs in once in a
browser window; after that every call is silent until the token is revoked.

**Delegated permissions**: the agent sees exactly what the signed-in person can
already see. There is no service account and no tenant-wide access — if they
cannot open a site, neither can you.

Runs on the interpreter Hermes already uses; `msal`, `requests`, `python-docx`,
`openpyxl`, `pymupdf` (`>=1.26.6,<2`), `pdfminer.six` and
`azure-ai-documentintelligence` are installed there. python-docx reads a
`.docx`, openpyxl reads a `.xlsx`. PyMuPDF reads a PDF text layer and any
table locally, and renders a scanned page. That image goes to `M365_OCR_URL` (default
`http://ihai.inhat.hu:11500/v1`, model `M365_OCR_MODEL`, default
`deepseek-ai/DeepSeek-OCR-2`). Set `M365_OCR_URL` empty to skip the internal
OCR. Azure Document Intelligence runs only when `AZURE_DOC_INTEL_ENDPOINT`
and `AZURE_DOC_INTEL_KEY` are set and the local page was empty. Azure vision
uses `AZURE_SWEDEN_API_KEY` and `gpt-5.6-luna-datazone` for pages that are
still empty. The Graph token is not sent with the file or the image.

Shorthand below: `mg` = `%LOCALAPPDATA%\hermes\hermes-agent\venv\Scripts\python.exe ${HERMES_SKILL_DIR}/scripts/microsoft_graph.py`

## Preflight — always first

```
mg status
```

Tells you whether anyone is signed in, as whom, and which scopes were granted.

- **`signed_in: false`** → run `mg login`. It opens a real browser window;
  **the user has to be at the keyboard.** Say so before you start it, and
  never run it inside a scheduled job.
- **`msal_installed: false`** → re-run `Install-Hermes.ps1`, or
  `python.exe -m pip install msal`.
- **`token_error`** → the cached token died. `mg login` again.

## Signing in

```
mg login                       # full read set, browser window opens
mg login --minimal             # only User.Read + Mail.Read
mg login --add-scope Calendars.ReadWrite
```

Consent is **incremental**: `--add-scope` keeps what was already granted and
asks for one more. Never widen the scope set on your own initiative — a scope
is a permission the user grants, so ask them first and tell them what for.

If consent is refused, the error names the AADSTS code and what it means. The
common ones: user consent disabled in the tenant (an admin must approve),
or a scope that always needs an admin — **Teams *channel* messages
(`ChannelMessage.Read.All`) are in that group**, which is why `mg search`
covers chats but not channel posts.

## One-shot examples

### Search everything

```
mg search --query "beszerzési eljárás 2026"
mg search --query "költségvetés" --types driveItem,listItem --top 25
mg search --query "from:someone@example.org contract" --types message
mg search --query "filetype:xlsx negyedéves" --types driveItem
mg search --query "határidő" --types chatMessage
```

Types: `message` (mail), `driveItem`/`listItem`/`site` (OneDrive+SharePoint),
`chatMessage` (Teams chats), `event`, `person`. Default is
`message,driveItem,chatMessage`. Graph refuses to mix most of these in one
request — the script fans them out and merges, so just ask for what you want.

The query is **KQL**: bare words, `from:`, `filetype:`, `subject:`,
`AND`/`OR`/`NOT`, `"exact phrase"`. It is keyword search, not semantic — if a
query returns nothing, try the words that would literally appear in the
document before concluding it does not exist.

Results are compacted to the readable fields plus an `id` and a `link`. Use
`--raw` only when you need a field that was dropped, and `--start` to page.

### Mail

```
mg mail list --folder inbox --top 20
mg mail get --id <id-from-list-or-search>
```

`search --types message` finds messages anywhere; `mail list` is for "what
came in recently".

### Calendar

```
mg calendar list --days 7
```

### OneDrive — list, enter, then read

The signed-in user's own drive, first:

```
mg get /me/drive/root/children
```

The answer is already counted. Report `count`, then the names in `folders`,
then the names in `files`. Do not add those lists up again. Enter a folder
with the `id` from `folders`:

```
mg get /me/drive/items/<id>/children
```

`/sites/root/drive/root/children` is the organisation's root site library,
not this user's OneDrive. Use it only when the user asked for that site.
Never call `get` without a path: Graph answers with its API catalogue,
which is not the user's files.

Read a Word, Excel or PDF file with the `id` from `files`:

```
mg get /me/drive/items/<id>/content
```

The answer is JSON with `saved_to`, `text` and `text_source`. Summarize
`text`. For a workbook, keep the sheet names. `text_source` is `docx` or
`xlsx`, or for a PDF the layers joined with `+`: `text-layer`, `local-vlm`,
`azure-di`, `pdfminer`, `azure-vision`. Do not call `/preview` and do not
open `@microsoft.graph.downloadUrl` or the SharePoint `webUrl` in a browser:
that session is not signed in, and the plugin would also misread the file
bytes as JSON. `/content` follows the redirect without the Graph token,
which is the only way the bytes come back.

Other reads:

```
mg get /me/drive/recent
mg get /me/joinedTeams
mg get /me/messages --param '$select=subject,receivedDateTime' --all
```

`get` reaches any Graph GET endpoint. Look the path up in the Microsoft Graph
documentation when you need something not listed here — the whole read surface
of Graph v1.0 is available this way.

### From `execute_code` (preferred for multi-step work)

One tool call does search → read → cross-reference, instead of three:

```python
import sys
sys.path.insert(0, r"${HERMES_SKILL_DIR}/scripts")
from microsoft_graph import graph, search, GraphError

hits = search("beszerzés 2026", types=["driveItem"], top=10)
for h in hits["hits"]:
    print(h["name"], h.get("location"), h.get("link"))

me = graph("GET", "/me")
graph("GET", "/me/messages", params={"$top": 5, "$select": "subject"})
```

`graph(method, path, params=, body=, allow_write=, all_pages=)` is the whole
API. No wrapper layer to learn: paths and query parameters are exactly what the
Graph documentation says.

## Writing — closed by default

`graph()` **refuses every non-GET call** unless the caller names a write area
and the path belongs to it. The gate is in the code, not in this document, so
it holds regardless of what you decide:

| Area | Covers | Scope needed |
|---|---|---|
| `calendar` | create/update/delete events on `/me/events` | `Calendars.ReadWrite` |
| `mail-send` | `/me/sendMail`, reply, forward | `Mail.Send` |
| `files` | `/me/drive/...` | `Files.ReadWrite.All` |

`POST /search/query` is exempt — it is the search API and changes nothing.
`/$batch` is allowed only when every sub-request is a GET.

**Before any write: say what you are about to change and get the user's
explicit agreement in the conversation.** A calendar event with attendees
sends invitations to real people the moment it is created, and that cannot be
taken back.

```
mg calendar add --subject "Egyeztetés" --start 2026-08-26T10:00:00 \
                --end 2026-08-26T11:00:00 --location "Nagy tárgyaló"
```

Add `--attendees a@b.hu,c@d.hu` only when the user asked for invitations by
name. If the scope is missing you get a message naming the exact
`login --add-scope` command — relay it, do not try to work around it.

## Data handling — read this before searching case files

Whatever you retrieve enters the model's context. With a locally hosted model
that stays inside your network. **With a hosted cloud model it leaves the
organisation.** If the user is on a cloud model and asks you to search M365
content, say so before you search and offer to switch to a local model first.

Quote what you found and link to it; do not bulk-copy documents into the
conversation when a summary and a link will do.

## Failure modes

| Symptom | Cause | What to do |
| --- | --- | --- |
| `not signed in to Microsoft 365` | no cached token | `mg login`, with the user present |
| `consent was not granted (AADSTS65001)` | user declined, or consent is disabled | try `--minimal`; if that fails too, an admin must approve |
| `needs administrator approval (AADSTS90094)` | admin-only scope requested | drop that scope; channel search is not available without an admin |
| `application is not present in this tenant (AADSTS700016)` | default client blocked | set `MSGRAPH_CLIENT_ID` to your own app registration |
| `registered single-tenant ... /organizations (AADSTS50194)` | own app is single-tenant, tenant id not set | set `MSGRAPH_TENANT_ID` next to `MSGRAPH_CLIENT_ID` |
| `does not allow public client flows (AADSTS7000218)` | registration setting | Entra: Authentication -> Allow public client flows = Yes; no secret needed |
| `multi-factor authentication is required` | MFA prompt | the user must complete it in the browser |
| `denied the request (403)` | missing scope, or genuinely no access | check `mg scopes` before assuming a bug |
| `partial_failures` in a search result | one entity group failed, others worked | read which scope is missing; the rest of the hits are still valid |
| `refused: ... is a write` | the write gate | name the area, or ask the user |
| Search returns nothing | KQL is literal | retry with words that appear verbatim; confirm the user has access |

## Boundaries

- **Never sign in on the user's behalf in an unattended context** — no
  scheduled jobs, no "I'll just log in and check". The browser window is their
  consent moment.
- **Never widen scopes silently.** Every new scope is a new permission over
  their mailbox or files.
- Do not send mail, delete anything, or invite people without an explicit
  request in the conversation. `mail-send` exists for when they ask; it is not
  a convenience.
- `mg logout` clears the local cache only. Revoking the consent itself is done
  by the user in their Microsoft account settings — tell them that if they ask
  how to withdraw access.
- This skill is for **cloud** content. A document open in Word on this machine
  belongs to `office-live`; a file on a network share is an ordinary file.
