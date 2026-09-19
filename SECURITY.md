# Security

## What these skills are allowed to do

Both run as **you**, with your privileges, on your machine. Neither asks an
operating system for anything you could not do yourself, and neither gains you
any access you did not already have. That is also the limit of the protection:
an agent driving them can reach every file, mailbox and share that you can.

`office-live` drives the Office applications already running on your desktop
through COM. It can read and change any document you have open. Word edits are
made as tracked revisions so they can be rejected, but Excel and PowerPoint
have no such mechanism and changes there are immediate.

`m365-graph` reads your own Microsoft 365 content with a delegated token: mail,
OneDrive and SharePoint files, Teams chats, calendar. It sees exactly what you
see, never more. Writes are refused in code rather than by convention — a
non-read call has to name a write area from a fixed list, and the request path
has to belong to that area, or the call is rejected before anything leaves the
machine.

## The part that will trip a security scanner

Both skills execute Python that a language model wrote. That is not incidental,
it is the design: a fixed tool list cannot cover the Office object model, so the
skill hands the agent the live objects and lets it write the four lines it
needs.

Concretely, `scripts/office_kernel.py` calls `exec(compile(...))`, and
`scripts/office.py` spawns a subprocess. Hermes Agent's skill-hub scanner
flags both and refuses to install these skills without `--force`, which is the
correct behaviour on its part: from the outside those patterns are
indistinguishable from a hostile skill. Trust here comes from reading the
source and from where you got it, not from a scanner's approval.

If that trade-off is not acceptable in your environment, do not install these.

## What does not happen

- **No outbound network from `office-live`.** Its only socket is a loopback
  listener (see below). It contacts nothing.
- **No telemetry, from either skill.** Nothing is reported anywhere.
- **No credential is stored by us.** `m365-graph` uses MSAL's own token cache;
  the sign-in happens in your browser and we never see the password.
- **Nothing is sent to a model by the skill itself.** Whatever the agent reads
  enters its context, and where that context goes is decided by the model you
  configured, not here. With a cloud model, content leaves your organisation.
  That is the single most important thing to understand before pointing these
  at case files.

## The warm daemon

`office.py --warm` keeps a background process holding the COM objects and the
Python namespace between calls, because cold-starting COM costs a third of a
second each time. It binds `127.0.0.1` on an ephemeral port, requires a
128-bit random token on every request, and writes the port and token to
`%LOCALAPPDATA%\office-live\daemon.json`, which is inside your own profile. It
exits by itself after 45 idle minutes, and restarts when it notices its own
source file has changed.

The token file is protected by the ACL on your user profile directory. On a
machine where another account can read your profile, another account can drive
your Office session. Treat a shared or roaming profile accordingly.

## The Microsoft Graph client id

By default `m365-graph` authenticates with the public client id of **Microsoft
Graph Command Line Tools** (`14d82eec-…`), which is Microsoft's own published,
well-known id. It is not a secret and it is not ours.

Using it has two consequences worth knowing:

- your tenant's sign-in log names that application, not this skill;
- tenant-wide consent granted to that id is granted to *anything* using it.

For anything beyond trying the skill out, register your own Entra ID
application and set `MSGRAPH_CLIENT_ID`. No code change is needed. The
requested scopes are then only yours, the sign-in log names your application,
and consent can be revoked without affecting anything else.

## Reporting a problem

Open an issue on the repository. If the problem is one you would rather not
describe in public, write to `mate.benyovszky@integritashatosag.hu` instead.

This is published work, not a supported product: there is no response-time
commitment and no security maintenance guarantee. See the disclaimer in the
README.
