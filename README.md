# Teams Proactive Message Test

Bare-minimum proof that an app (no signed-in user) can send a Teams message
directly to a user. This is the same pattern your MCP server would use.

**Status: proven working end to end** — bot receives messages through a dev
tunnel, captures conversation references, and pushes proactive 1:1 messages
using only app credentials.

## Why a bot, and not Graph or a webhook?

- **Graph API, app-only:** cannot create 1:1 chat messages. Sending chat
  messages via Graph is delegated-only (a real user must be signed in).
- **Incoming webhooks:** channel-only, can't DM a user, and the underlying
  Office 365 connectors are retired.
- **Azure Bot + proactive messaging:** the supported way for a headless
  service to message a user. The bot needs a *conversation reference*
  (conversation id + service URL), which it captures the first time a user
  installs or messages it. After that, the app can push messages any time.

No Bot Framework SDK is used here — just two REST calls (get a token, post an
activity) — so the logic ports directly into any MCP server in any language.

## Privileges required — the least-privilege case

This is the complete inventory. There is nothing else.

| Privilege | Scope | Who grants it |
|---|---|---|
| App registration + credential (client secret, or a federated credential — see "Hosting in Azure" below for the secretless option) | Identity only — grants nothing by itself | Entra ID (any app admin) |
| Service principal for that app | Required for the tenant to issue tokens to the app (`az ad sp create --id <appId>`) | Entra ID |
| Azure Bot resource (F0, free) | Control plane only — links the app ID to the Teams channel | Azure subscription contributor |
| Teams app in the org catalog | Lets users (or admins) install the bot | Teams admin |

**Microsoft Graph API permissions required: none.** The app's token is scoped
to `https://api.botframework.com/.default` (the Bot Framework Connector), not
Graph. There is no admin-consent prompt because there is nothing to consent
to — the app registration has an empty API-permissions list.

### What the bot identity can and cannot do

| Can | Cannot |
|---|---|
| Send messages into conversations where its Teams app is installed | Message any user who hasn't got the app installed — Teams rejects it (`403 Forbidden`) |
| Receive messages users deliberately send *to it* | Read any other chat, channel, or message — it only sees what is addressed to it |
| | Access mail, calendars, files, or the directory — no Graph permissions exist |
| | Impersonate a user — messages are visibly from the bot, with the bot's name and icon |

The key property: **the recipient boundary is enforced by Microsoft, not by
our code.** The bot can only reach conversations it is a member of, and
membership is controlled by Teams app installation, which the Teams admin
governs centrally.

### Compared to the alternatives IT might fear

| Approach | Blast radius |
|---|---|
| **This (bot, proactive messaging)** | Send-only, to app-installed users only, no read access to anything |
| Graph `Chat.ReadWrite.All` (application) | Read **and write** every chat of every user in the tenant; protected API requiring Microsoft approval |
| Service account + delegated Graph | A licensed user account with a credential to protect; can read/send as that account everywhere it has access |
| Exchange `Mail.Send` (application) | Send mail as *any* user unless separately restricted with an application access policy |

## Restricting who can receive messages

Two layers — one enforced by the platform, one by the app.

### 1. Platform enforcement (the hard boundary)

The bot physically cannot message a user whose Teams client doesn't have the
app installed. So recipient scope **is** app-availability scope, managed in
the Teams admin center:

- **Teams admin center → Teams apps → Manage apps → \<your app\> → Users and
  groups**: make the app available to specific users or security groups only
  (e.g., a `developers` group). Everyone else cannot install it, therefore
  cannot be messaged. This is auditable, centrally controlled, and requires
  no code.
- Removing a user from the group / uninstalling the app immediately revokes
  the bot's ability to message them.

### 2. Application-level allowlist (defense in depth)

Check the sender/recipient against an allowlist before storing references or
sending. Every Teams activity carries the user's Entra object ID
(`from.aadObjectId`), so the check is one line:

```js
const ALLOWED = new Set((process.env.ALLOWED_USER_IDS || "").split(","));
if (!ALLOWED.has(aadObjectId)) return res.status(403).json({ error: "recipient not allowed" });
```

For the real MCP server, resolve an Entra **group membership** instead of a
static list, so IT owns the roster in one place (this is the one case where
a single Graph permission — `GroupMember.Read.All` — could be added, and it
is read-only).

## Zero-touch onboarding: auto-install the app for an Entra group

Users don't have to install the bot themselves. A Teams **app setup policy**
can install it for them, targeted by group:

1. Teams admin center → Teams apps → **Setup policies** → create a policy →
   under **Installed apps**, add the bot app.
2. Assign the policy to an Entra security group (Group policy assignment
   tab) — e.g. a `developers` group.

Everyone in the group gets the app silently installed into their Teams
client. Two properties matter here:

- **Zero additional privileges.** The roster is an admin-controlled policy +
  Entra group. The app's permission list stays empty; the service never
  touches the directory. IT owns who is in scope.
- **No user action needed.** When Teams auto-installs the app, the bot
  endpoint receives a `conversationUpdate`/`installationUpdate` activity, and
  this server already captures the conversation reference from those events —
  not just from messages. Users become messageable without ever opening the
  bot.

Combined with the availability scoping above, the same group can drive both:
who *can* have the app (Manage apps → Users and groups) and who *gets it
automatically* (setup policy). Caveat: policy assignment isn't instant — it
can take up to 24 hours to propagate to clients.

(The code-based alternative — the service installs the app itself via Graph
with `TeamsAppInstallation.ReadWriteSelfForUser.All` — is only needed if
onboarding must happen on demand, faster than policy sync.)

## Limiting how many messages can be sent

### Platform throttling (always on)

Teams rate-limits bots service-side. Exceeding per-conversation or per-tenant
thresholds returns `HTTP 429` with a `Retry-After` header — the platform will
not let a runaway bot flood the tenant. Microsoft's guidance is roughly one
message per second per conversation sustained; details:
<https://learn.microsoft.com/microsoftteams/platform/bots/how-to/rate-limit>

### Application-level caps (policy you control)

The platform limit is an abuse backstop, not a policy tool. Enforce your own
budget in the send path, e.g. a per-user daily cap:

```js
const sentToday = {}; // userId -> count, reset daily (use real storage in prod)
const DAILY_CAP = 10;
if ((sentToday[aadObjectId] || 0) >= DAILY_CAP) {
  return res.status(429).json({ error: "daily message cap reached for this user" });
}
sentToday[aadObjectId] = (sentToday[aadObjectId] || 0) + 1;
```

Sensible production shape: per-user cap (protects individuals from noise) +
global cap (bounds total volume) + an audit log of every send (who, when,
what, triggered by which MCP tool call). The audit log is cheap and is
usually what security actually asks for.

## What you need

- An M365 tenant where you're admin (your personal/dev tenant) with Teams.
- **Custom app upload enabled:** Teams admin center → Teams apps → Setup
  policies → Global → turn on "Upload custom apps".
- An Azure subscription (the Azure Bot resource is free at F0 tier).
- Node 18+.

## Setup

### 1. App registration

Entra ID → App registrations → New registration:
- Name: anything (e.g. `teams-notify-test`)
- Supported account types: **Accounts in this organizational directory only** (single tenant)
- No redirect URI needed.

Then: Certificates & secrets → New client secret. Record:
- Application (client) ID → `BOT_APP_ID`
- Secret value → `BOT_APP_SECRET`
- Directory (tenant) ID → `TENANT_ID`

**Also create the service principal** — the portal does this implicitly in
some flows, the CLI does not, and without it token requests fail with
`AADSTS7000229`:

```powershell
az ad sp create --id <BOT_APP_ID>
```

No API permissions are needed.

CLI equivalent of the whole step:

```powershell
az ad app create --display-name teams-notify-test --sign-in-audience AzureADMyOrg
az ad sp create --id <appId>
az ad app credential reset --id <appId> --display-name bot-secret --years 1
```

### 2. Azure Bot resource

Azure portal → Create resource → **Azure Bot**:
- Pricing tier: Free (F0)
- Type of App: **Single Tenant**
- Creation type: **Use existing app registration** → paste the app ID and tenant ID from step 1.

After it deploys:
- Settings → Channels → add **Microsoft Teams**.
- Settings → Configuration → set **Messaging endpoint** to your public URL +
  `/api/messages` (see step 4).

CLI equivalent:

```powershell
az group create --name rg-teams-notify-test --location eastus
az bot create --resource-group rg-teams-notify-test --name <botName> `
  --app-type SingleTenant --appid <appId> --tenant-id <tenantId> `
  --sku F0 --endpoint https://placeholder.example.com/api/messages
az bot msteams create --name <botName> --resource-group rg-teams-notify-test
```

### 3. Run the app

```powershell
Copy-Item .env.example .env   # then fill in the three values
npm install
npm start
```

### 4. Expose it publicly

Teams must be able to reach `/api/messages`. Easiest local option is the
Microsoft dev tunnel CLI:

```powershell
winget install Microsoft.devtunnel
devtunnel user login
devtunnel create teams-bot-test --allow-anonymous
devtunnel port create teams-bot-test -p 3000
devtunnel host teams-bot-test
```

Copy the public URL it prints and update the bot's messaging endpoint:

```powershell
az bot update --resource-group rg-teams-notify-test --name <botName> `
  --endpoint https://<tunnel-url>/api/messages
```

### 5. Teams app package

1. Edit `manifest/manifest.json`: set `id` and `bots[0].botId` to your
   `BOT_APP_ID`.
2. Zip the three files (must be at the zip root, not in a folder):

```powershell
Compress-Archive -Path manifest\manifest.json, manifest\color.png, manifest\outline.png -DestinationPath notify-test-bot.zip -Force
```

3. In Teams: Apps → Manage your apps → **Upload an app** → "Upload a custom
   app" → pick the zip.

### 6. Test

1. In Teams, open the installed bot and send it any message ("hi"). The bot
   echoes back a confirmation and the server stores the conversation
   reference in `conversations.json`.
2. Open http://localhost:3000, pick the user, type a message, click send.
3. The message arrives in Teams as a 1:1 chat from the bot — no user session
   involved, pure app credentials.

## Hosting in Azure (free) with federated credentials — no client secret

The whole stack also runs at $0/month on Azure, and doing so removes the
client secret entirely: a **user-assigned managed identity** plus a
**federated credential** on the app registration replaces it. There is then
no secret to rotate, leak, or expire anywhere.

How it works: `getBotToken()` picks the flow automatically. If
`BOT_APP_SECRET` is set (local dev) it uses the classic secret flow.
Otherwise it requests a managed-identity token for the special audience
`api://AzureADTokenExchange` from the App Service identity endpoint and
presents it as a `client_assertion` to the same Entra token endpoint —
two REST calls, still no SDK.

```powershell
# 1. Managed identity (free)
az identity create --name id-teams-notify-test --resource-group rg-teams-notify-test --location eastus
# note its clientId and principalId

# 2. Federated credential on the app registration, trusting that identity
#    (fic.json: name, issuer https://login.microsoftonline.com/<tenantId>/v2.0,
#     subject = identity principalId, audiences ["api://AzureADTokenExchange"])
az ad app federated-credential create --id <BOT_APP_ID> --parameters fic.json

# 3. Free App Service with the identity attached (F1 quota is per-region — try
#    another region if you get a quota error)
az appservice plan create --name plan-teams-notify-test --resource-group rg-teams-notify-test `
  --location westus2 --sku F1 --is-linux
az webapp create --name <appName> --plan plan-teams-notify-test `
  --resource-group rg-teams-notify-test --runtime "NODE:22-lts"
az webapp identity assign --name <appName> --resource-group rg-teams-notify-test --identities <identity resource id>
az webapp config appsettings set --name <appName> --resource-group rg-teams-notify-test --settings `
  BOT_APP_ID=<BOT_APP_ID> TENANT_ID=<TENANT_ID> AZURE_CLIENT_ID=<identity clientId> `
  SCM_DO_BUILD_DURING_DEPLOYMENT=true   # note: no secret

# 4. Deploy (use tar.exe, not Compress-Archive — the latter writes backslash
#    paths into the zip, which breaks extraction on Linux)
tar.exe -a -cf deploy.zip server.js package.json package-lock.json public
az webapp deploy --name <appName> --resource-group rg-teams-notify-test --src-path deploy.zip --type zip

# 5. Teams now reaches the bot directly — no dev tunnel
az bot update --resource-group rg-teams-notify-test --name <botName> `
  --endpoint https://<appName>.azurewebsites.net/api/messages
```

F1 caveats: no Always On, so the app sleeps after ~20 minutes idle and the
first *inbound* message after that hits a cold start (a few seconds —
proactive sends are unaffected); 60 CPU-minutes/day. Also note a redeploy
overwrites `conversations.json` in `wwwroot` — references are recaptured on
the next message, or use real storage.

## What this means for the real MCP server

- The conversation reference only needs to be captured **once per user**, then
  it's valid indefinitely — persist it in real storage.
- Users can be onboarded without ever messaging the bot: assign an app setup
  policy to an Entra group (see "Zero-touch onboarding" above) and the
  conversation references arrive automatically as Teams installs the app.
- What you'd need from central IT: an app registration (with a secret, or
  secretless via a federated credential + managed identity) + service
  principal, an Azure Bot resource, and the Teams app approved into the org
  catalog scoped to a target group. That's the full list — no mailbox, no
  Graph chat permissions, no admin consent.
- Production hardening not in this sample: validate the JWT on incoming
  `/api/messages` requests (issuer `api.botframework.com`), cache the bot
  token (it's good for ~1 hour), store references somewhere durable, and add
  the allowlist + caps + audit log described above.
