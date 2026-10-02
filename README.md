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

Both are Windows-and-Microsoft specific, which is why they live here rather
than upstream in any one agent's repository.

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

## As Hermes plugins

The same code is also packaged as two [Hermes Agent](https://hermes-agent.nousresearch.com)
plugins, so a colleague installs one thing instead of copying folders and
editing `config.yaml`:

| Plugin | What it gives Hermes | Tools |
|---|---|---|
| `office-connector` | Work in the Word, Excel and PowerPoint documents you already have open | `office_status`, `office_exec`, `office_reset` |
| `ms-graph-connector` | Read your own Microsoft 365: Outlook mail, OneDrive and SharePoint files, Teams chats, calendar | `m365_status`, `m365_login`, `m365_search`, `m365_get` |

```bash
hermes plugins install integrityauthority/office-agent-skills#plugins/office-connector
hermes plugins install integrityauthority/office-agent-skills#plugins/ms-graph-connector
```

Both pass `hermes plugins validate`, the catalogue admission gate.
`ms-graph-connector` scans clean; `office-connector` passes with a caution,
for the reason set out under **What they do and do not do**.

### One source, two shapes

The Python lives in exactly one place, `.agents/skills/<skill>/`. A plugin
directory is an adapter — a manifest and a `register()` — plus a **copy** of
that skill's files.

The copy is not a preference. `hermes plugins install` installs one directory
and nothing outside it comes along, so a plugin that reached up into
`.agents/` would validate here and be broken for everyone who installed it.
What keeps the copy honest is that it is generated:

```bash
py tools/sync_plugins.py            # refresh
py tools/sync_plugins.py --check    # fail on drift — CI runs this
```

Edit the skill, run the sync, commit both. Never edit a file under
`plugins/*/scripts/`.

### Plugin or skill?

Install the **plugin** if you use Hermes: the tools appear in the tool
catalogue where the model can see them, one warm COM connection serves the
whole session, and there is no MCP server to wire up.

Take the **skill** for anything else — Claude Code, Cursor, Codex — or when
you want the instructions without the tools. Note that a skill registered by
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
its context, and where that goes is decided by the model you configured.

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
[pywin32](https://github.com/mhammond/pywin32), and
[msal](https://github.com/AzureAD/microsoft-authentication-library-for-python)
(MIT), [requests](https://github.com/psf/requests) (Apache-2.0) and the
[MCP Python SDK](https://github.com/modelcontextprotocol/python-sdk) (MIT),
each installed and licensed on its own terms. No code from any other project
has been copied into this repository.

### Trademarks

Microsoft, Windows, Microsoft 365, Office, Word, Excel, PowerPoint, Outlook,
SharePoint, OneDrive, Teams and Microsoft Graph are trademarks of the Microsoft
group of companies. Hermes Agent is a project of Nous Research. They are named
here only to say what this software works with. No affiliation with, sponsorship
by or endorsement from any of them is claimed.
