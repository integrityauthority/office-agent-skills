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
> oversight; [SECURITY.md](SECURITY.md) explains the trade-off and what the
> skills deliberately do *not* do.

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
