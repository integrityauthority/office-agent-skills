# office-agent-skills

Agent Skills for Microsoft Office and Microsoft 365, in the portable
[Agent Skills](https://agentskills.io/specification) format. Each one is a
`SKILL.md` plus a thin Python layer the instructions tell the agent to import.

They run wherever Agent Skills run — [Hermes Agent](https://hermes-agent.nousresearch.com),
Claude Code, Cursor, Codex — and they are published through
[AgentPlaybooks](https://agentplaybooks.ai) so an agent can install and update
them from a URL.

| Skill | What it does | Needs |
|---|---|---|
| [`office-live`](.agents/skills/office-live) | Reads and edits the Word, Excel or PowerPoint document the user has **open**, over COM. Word edits land as tracked revisions. | Windows, Office, `pywin32` |
| [`m365-graph`](.agents/skills/m365-graph) | Searches the signed-in user's own Outlook mail, OneDrive and SharePoint documents, Teams chats and calendar, through the Microsoft Graph REST API. Read-only by default. | `msal`, a work or school account |
| [`file-reader`](.agents/skills/file-reader) | Reads files on disk in one call: PDF (scanned pages through Windows OCR, or an OCR server you configure), Outlook `.msg` and `.eml` e-mails with their attachments, Word, Excel, PowerPoint, HTML, text. Read-only. | `pypdfium2`, `olefile`; any OS |

The first two are Windows-and-Microsoft specific, which is why they live here
rather than upstream in any one agent's repository.

They follow one rule between them: **read in the background, write in front
of the user.** Reading documents — one or a folder of them, to answer,
search, summarise or collect — is file-level work with `file-reader` (or
`m365-graph` for the cloud) and opens nothing on screen. Producing a document
the user will keep happens with `office-live`, in a Word, Excel or PowerPoint
window open in front of them, so they see it take shape and can correct it.

> ### Read this before pointing them at real work
>
> These skills act as **you**, with your privileges. `office-live` can change
> any document you have open, and Excel and PowerPoint edits are immediate —
> only Word edits arrive as tracked revisions you can reject. `m365-graph`
> reads your own mail, files, chats and calendar.
>
> **Whatever the agent reads enters the model's context.** With a locally
> hosted model that stays inside your network; with a hosted cloud model it
> leaves your organisation. Decide that before you use these on anything
> confidential — the skills cannot decide it for you.
>
> Both skills run Python that the model wrote. That is the design, not an
> oversight — see **What they do and do not do** below.

## Why a thin Python layer instead of an MCP server

An MCP server for Office would need a fixed tool list, and the interesting part
of Office automation is exactly the part no fixed tool list covers. `office-live`
hands the agent the live COM object model plus two helpers — `explore()` to list
what an object really offers, and `probe()` to diff a document before and after
an edit — and lets it write the four lines it actually needs.

The same shape suits Graph better still: every call goes through one HTTP
function, so read-only can be enforced in code rather than by convention.

## Layout

```
.agents/skills/<name>/SKILL.md          the instructions the agent reads
.agents/skills/<name>/scripts/*.py      the module those instructions import
```

`.agents/skills/` is the vendor-neutral store. Grok Bot and Google Antigravity
read it natively, Hermes reads it through `skills.external_dirs`, and the
AgentPlaybooks CLI treats it as the portable source. Cloning this repository and
pointing an agent at it is therefore a complete installation — no build step,
no packaging.

## Using them

**Hermes Agent**, from a published playbook — no clone, no sign-up:

```bash
hermes skills install well-known:https://agentplaybooks.ai/.well-known/skills/office-live
```

**Any agent**, from a clone: copy `.agents/skills/<name>/` into that agent's
skill directory, or add this repository's `.agents/skills` to its search path.
For Hermes that is one line in `~/.hermes/config.yaml`:

```yaml
skills:
  external_dirs:
    - D:/office-agent-skills/.agents/skills
```

**Through a playbook**, which is what keeps a team in step:

```bash
apb pull <playbook-guid> --apply      # writes .agents/skills/, scripts included
```

## As plugins, for every platform

The same code is packaged as three plugins. Each plugin directory is one
package that several platforms read at once, so the install is one command
wherever you are:

| Plugin | What it gives the agent | Hermes tools |
|---|---|---|
| `office-connector` | Work in the Word, Excel, PowerPoint and Outlook windows you already have open (Windows) | `office_status`, `office_exec`, `office_reset` |
| `ms-graph-connector` | Read your own Microsoft 365: Outlook mail, OneDrive and SharePoint files, Teams chats, calendar | `m365_status`, `m365_login`, `m365_search`, `m365_get` |
| `file-connector` | Read and search PDF and images (with OCR for scanned pages and pictures), .msg and .eml e-mails with attachments, Word, Excel and PowerPoint files and .zip archives from disk; tabulate a folder of letters or documents | `file_read`, `file_find`, `file_collect`, `file_unpack`, `file_attachments`, `file_render` |

**Hermes Agent**: the tools appear in the tool catalogue.

```bash
hermes plugins install integrityauthority/office-agent-skills#plugins/file-connector
```

**Claude Code**: this repository is a plugin marketplace.

```bash
/plugin marketplace add integrityauthority/office-agent-skills
/plugin install file-connector@office-agent-skills
```

**Codex, Cursor, GitHub Copilot, VS Code and other
[Agent Plugins](https://vercel.com/kb/guide/agent-plugins) clients**: each
`plugins/<plugin>/` directory is an Agent Plugins 1.0 package (`plugin.json`
plus `skills/`). Install it the way your client installs a plugin directory.

**Any agent that reads [Agent Skills](https://agentskills.io)**: the skills
themselves are in `.agents/skills/`, and inside each plugin under `skills/`.

Every package passes its platform's own validator: `hermes plugins validate`
(security scan: safe for `file-connector` and `ms-graph-connector`;
`office-connector` passes with a caution, for the reason set out under
**What they do and do not do**), `claude plugin validate --strict`, and the
Agent Plugins 1.0 loader. The tests in `tests/test_packaging.py` hold the
packages to the same rules, so a store rejection is found in CI first.

### One source, every shape

The Python lives in exactly one place, `.agents/skills/<skill>/` (and
`shared/` for code two skills need). A plugin directory is generated around
it:

```
plugins/<plugin>/
  plugin.yaml, __init__.py      Hermes (hand-written: the tool adapter)
  plugin.json                   Agent Plugins 1.0
  .claude-plugin/plugin.json    Claude Code
  skills/<skill>/               the skill, copied: SKILL.md and scripts/
.claude-plugin/marketplace.json the Claude Code marketplace
```

Hermes reads `plugin.yaml` and ignores the `plugin.json` beside it, so the
manifests never compete. The copy is not a preference: every installer takes
one directory and nothing outside it, so a plugin that reached up into
`.agents/` would validate here and be broken for everyone who installed it.
What keeps the copies and the manifests honest is that they are generated
from the skill and `plugin.yaml`:

```bash
py tools/sync_plugins.py            # refresh
py tools/sync_plugins.py --check    # fail on drift — CI runs this
```

Edit the skill or `plugin.yaml`, run the sync, commit all of it. Never edit a
file under `plugins/*/skills/` or a generated `plugin.json`.

### Plugin or skill?

Install the **plugin** if you use Hermes: the tools appear in the tool
catalogue where the model can see them, one warm COM connection serves the
whole session, and there is no MCP server to wire up.

Take the **skill** (or the plugin, which carries it) for anything else —
Claude Code, Cursor, Codex — or when you want the instructions without the
tools; there the agent runs the same scripts from its terminal. Note that a skill registered by
a plugin is deliberately left out of Hermes' `<available_skills>`, so inside
Hermes the tools are what the agent discovers; the skill document rides along
as reference material.

## Developing a skill

Edit the files in place and run the checker:

```bash
py tools/check_skills.py
```

It asserts the things that are only discovered later otherwise: that the
frontmatter carries the `name` and `description` the specification requires and
that the name matches the directory, that every bundled file has a name and a
size the publishing path accepts, that the Python parses, and that every
`scripts/...` path the `SKILL.md` mentions is a file the skill actually ships.
That last one is the failure this repository exists to avoid — instructions
pointing at a module that never arrived.

CI runs the same command on every push, so a skill that cannot be published
cannot be merged.

Code more than one skill needs lives once, under `shared/` (today
`shared/pdfread.py`, the PDF and OCR reader of file-reader and m365-graph).
A skill is installed as one directory, so `tools/sync_plugins.py` copies the
module into each skill's `scripts/` and each plugin; CI fails if a copy drifts.
Edit `shared/`, never a copy.

The tests build every file they read (PDF, `.msg`, `.eml`, Office) and run a
fake OCR server, so they touch no real document, mailbox or model:

```bash
pip install -r tests/requirements.txt
py -m pytest
```

Before a push, `py tools/check_public.py --commits origin/main..HEAD` checks
your commits the way CI checks a pull request.

Write the `SKILL.md` for the agent, not for a person. An instruction phrased as
advice to a human reader ("export your token first") is read as an instruction
to the agent, which will then ask the user for something it already has.

## Publishing

A playbook is the distribution channel. Link this repository to one, push, and
the skills plus their Python files are served as plain Markdown and text over
HTTP:

```bash
apb login                             # once per machine
apb pull <playbook-guid> --apply      # links this checkout to that playbook
apb push --yes                        # SKILL.md and scripts/*.py go up
```

The `pull` is what writes `.agentplaybooks/remote.json`, and every later `push`
reads the target from there. Without it `push` offers to create a new playbook
instead, which is a different thing and will say so in its plan. Nothing is
uploaded until the plan is accepted.

Set the playbook's visibility to **public** and every skill in it appears at

```
https://agentplaybooks.ai/.well-known/skills/index.json
https://agentplaybooks.ai/.well-known/skills/<name>/SKILL.md
https://agentplaybooks.ai/.well-known/skills/<name>/scripts/<file>.py
```

which is the well-known convention Hermes installs from directly. There is no
registry to submit to and nobody to ask.

Updating is the same command. `apb push` adds files that are new and updates
files that changed; a file that exists in the playbook but no longer here is
left alone rather than deleted, so a push can never silently remove something
another maintainer added. On the consuming side, `hermes skills update` or a
fresh `apb pull` takes the new version.

Keep a playbook private while a skill is still being shaped: a private playbook
is reachable with `apb pull` and a key, and serves nothing publicly.

## Configuration

Everything is optional and read from environment variables; Hermes reads them
from its `.env`. Nothing in this repository names a server or a deployment.
With no configuration at all, a scanned page is still read on Windows, by
Windows' own OCR on the machine.

| Variable | Used by | Meaning |
|---|---|---|
| `OCR_LOCAL_LANGUAGE` | file-reader, m365-graph | Language for Windows OCR, such as `hu` or `en-US`. Default: the user's Windows languages. It must be installed (Settings → Time & language → Language). |
| `OCR_LOCAL` | file-reader, m365-graph | `off` turns Windows OCR off. |
| `OCR_URL` | file-reader, m365-graph | Your own OpenAI-compatible OCR server (`/chat/completions` with an image), for when Windows OCR is not good enough. |
| `OCR_MODEL` | file-reader, m365-graph | Model on that server. Default `deepseek-ai/DeepSeek-OCR-2`; any other vision model gets a plain transcription prompt. |
| `OCR_API_KEY` | file-reader, m365-graph | Bearer token for that server, if it wants one. |
| `OCR_CLOUD_URL` | file-reader, m365-graph | A cloud vision model behind the OpenAI Responses API (`/responses`), Azure OpenAI included. |
| `OCR_CLOUD_MODEL` | file-reader, m365-graph | Model or deployment name there. |
| `OCR_CLOUD_API_KEY` | file-reader, m365-graph | Its key. Sent as `api-key` to `*.azure.com`, as a bearer token elsewhere. |
| `OCR_CLOUD_API_KEY_ENV` | file-reader, m365-graph | Instead of the key: the name of the variable that already holds it, so a secret your agent has anyway is not copied. |
| `ZIP_NAME_ENCODING` | file-reader | Code page of `.zip` member names stored without the UTF-8 flag (Windows Explorer and older archivers write the DOS code page). Default `auto`: UTF-8 when the bytes are valid UTF-8, otherwise whichever of `cp852`, `cp1250` and `cp437` reads as letters. Set one code page, such as `cp852`, to use only that one. |
| `MSGRAPH_CLIENT_ID`, `MSGRAPH_TENANT_ID` | m365-graph | Your own Entra ID application and tenant; see below. |

The OCR order is the same in both skills, cheapest and most private first:

1. the text layer, on this machine, always;
2. a scanned page: **Windows OCR on this machine** (`ocr="auto"`, the
   default; well under a second a page, nothing leaves the computer). Where
   there is no Windows OCR, your OCR server instead;
3. **your OCR server** when the agent or the user finds that text not good
   enough: `ocr="server"`. It keeps table structure and reads numbers more
   reliably. The answer says which pages Windows read (`ocr_upgrade`), so the
   agent knows when to ask for it;
4. **a cloud model only on a yes**: when pages are still empty, the answer
   carries `cloud_ocr_needs_approval`; the agent asks the user and only then
   calls again with `ocr="cloud"`.

Windows OCR is reached through [pywinrt](https://github.com/pywinrt/pywinrt)
(MIT) in-process, with no helper program started. On another operating
system step 2 is skipped. There is no Azure Document Intelligence step: it
needed its own SDK and sent the whole PDF, not only the empty pages.

## What they do and do not do

Both skills act as the signed-in user and gain no access that user did not
already have. `office-live` drives the Office applications already running on
the desktop through COM; Word edits are tracked revisions, Excel and
PowerPoint edits are immediate. `m365-graph` reads the user's own Microsoft
365 content with a delegated token and refuses non-read calls in code — a
write has to name an area from a fixed list and the request path has to belong
to it, or the call never leaves the machine.

What does not happen: `office-live` makes **no outbound network calls at all**,
neither skill reports telemetry anywhere, and no credential is stored by them
(`m365-graph` uses MSAL's own token cache and never sees a password). Nothing
is sent to a model by the skills themselves — whatever the agent reads enters
its context, and where that goes is decided by the model you configured. The
one exception is OCR, described under [Configuration](#configuration): a
scanned page is read by Windows OCR on the machine, goes to your own OCR
server only when the agent asks for better text, and to a cloud model only on
a second, explicit call the user has agreed to. A page with a text layer is
always read locally.

`office.py --warm` keeps a background process holding the COM objects between
calls, because cold-starting COM costs a third of a second each time. It binds
`127.0.0.1` on an ephemeral port, requires a 128-bit random token on every
request, and exits by itself after 45 idle minutes. The token lives in the
user's own profile directory, so a shared or roaming profile deserves thought.

**The part that trips a security scanner.** `scripts/office_kernel.py` calls
`exec(compile(...))` and `scripts/office.py` spawns a subprocess. Both are
deliberate: a fixed tool list cannot cover the Office object model, so the
skill hands the agent the live objects and lets it write the few lines it
needs. From the outside those patterns are indistinguishable from a hostile
skill, which is why Hermes Agent's skill-hub scanner refuses to install these
without `--force` — correctly. Trust here comes from reading the source and
from knowing where you got it, not from a scan. If that trade-off does not
work in your environment, do not install these.

**The Microsoft Graph client id.** By default `m365-graph` authenticates with
the public client id of Microsoft Graph Command Line Tools, Microsoft's own
published id. It is not a secret and not ours, but it has two consequences:
your tenant's sign-in log names that application rather than this skill, and
tenant-wide consent granted to that id is granted to anything using it. For
anything beyond trying the skill out, register your own Entra ID application
and set `MSGRAPH_CLIENT_ID`. No code change is needed.

## Licence, warranty, and what this is not

MIT — see [LICENSE](LICENSE). Use, modify and redistribute it freely,
commercially or otherwise, provided the copyright notice and the licence text
travel with it.

**Provided as is, without warranty of any kind.** That is the MIT text and it
is meant literally here. This is working code published because it may be
useful to others, not a product:

- no fitness for any particular purpose is claimed or implied;
- no support, no service level, no security maintenance commitment;
- no liability for data loss, for a document edited in a way you did not
  intend, or for anything an agent does while driving these skills;
- publishing it is not advice, endorsement, certification or approval of any
  kind by Integritás Hatóság, and it says nothing about the authority's
  official activity.

Test it on copies before you let it near anything that matters.

### If you deploy this

You become the operator. The skills read personal data — mail, documents,
chats, calendar — and pass it to whichever model you configure. Deciding
whether that is lawful in your setting, on what basis, and with what
safeguards, is yours: these skills have no telemetry and send nothing anywhere
by themselves, so every flow of data out of your environment is one you
configured. `m365-graph` reads only what the signed-in user can already read,
with delegated consent, and refuses non-read calls unless they name a write
area explicitly.

### Third-party software

Nothing third-party is bundled here. At runtime the skills import
[pywin32](https://github.com/mhammond/pywin32) (PSF),
[msal](https://github.com/AzureAD/microsoft-authentication-library-for-python)
(MIT), [requests](https://github.com/psf/requests) (Apache-2.0),
[pypdfium2](https://github.com/pypdfium2-team/pypdfium2) (Apache-2.0 / BSD-3),
[pdfplumber](https://github.com/jsvine/pdfplumber) and
[pdfminer.six](https://github.com/pdfminer/pdfminer.six) (MIT),
[olefile](https://github.com/decalage2/olefile) (BSD-2),
[openpyxl](https://foss.heptapod.net/openpyxl/openpyxl) (MIT),
[xlrd](https://github.com/python-excel/xlrd) (BSD-3),
[python-docx](https://github.com/python-openxml/python-docx) (MIT) and
[Pillow](https://github.com/python-pillow/Pillow) (MIT-CMU),
[pywinrt](https://github.com/pywinrt/pywinrt) (MIT, Windows only), and optionally
the [MCP Python SDK](https://github.com/modelcontextprotocol/python-sdk) (MIT),
each installed and licensed on its own terms.

Only permissively licensed dependencies are used, on purpose. PyMuPDF
(AGPL-3.0) and extract-msg (GPL-3.0) would each do part of the job, and are
not used: the `.msg` reader here is written against Microsoft's published
[MS-OXMSG] format instead.

[MS-OXMSG]: https://learn.microsoft.com/openspecs/exchange_server_protocols/ms-oxmsg/ No code from any other project
has been copied into this repository.

### Trademarks

Microsoft, Windows, Microsoft 365, Office, Word, Excel, PowerPoint, Outlook,
SharePoint, OneDrive, Teams and Microsoft Graph are trademarks of the Microsoft
group of companies. Hermes Agent is a project of Nous Research. They are named
here only to say what this software works with. No affiliation with, sponsorship
by or endorsement from any of them is claimed.
