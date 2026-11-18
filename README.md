# Council

One desktop window that asks **Claude, Gemini and Copilot** the same question, shows the three answers side by side, and can have one of them combine the answers into a single best one.

There are no API keys and nothing is billed per use. Each model runs through its own official command-line tool, signed in with the subscription you already have:

| Pane | Tool | Signs in with |
|---|---|---|
| Claude | Claude Code (`claude`) | Claude Pro |
| Gemini | Antigravity CLI (`agy`), Google's replacement for Gemini CLI | Google AI Pro |
| GitHub Copilot | GitHub Copilot CLI (`copilot`) | GitHub Copilot (the Free plan works) |
| M365 Copilot | Microsoft 365 Copilot Chat API | Microsoft 365 Business plan with Copilot |
| Local | Ollama on this PC | Nothing: free, private, works offline |

## The Local pane

- It uses a model you've downloaded in Ollama. Pick which one from the dropdown at the top of the pane; Council remembers your choice.
- Ollama has to be running (Start menu > Ollama). If the badge says "Ollama not running", start it and click the badge.
- To add a model, open a terminal and run `ollama pull <name>`, for example `ollama pull gpt-oss:20b` (about 14 GB). Then click the Local badge so the dropdown picks it up.
- Bigger models answer better but take longer and use more graphics-card memory. The model is unloaded 5 minutes after your last question (`keep_alive` in Settings).
- If every online model is unticked or not signed in, Council still works with just the Local pane.
- Desktop shortcut: **Council** on your desktop runs `run_council.bat`.

## Setup (once)

1. Install **Python 3** (you already have it) and **Node.js 20+**.
2. Double-click **setup.bat**. It installs all three tools.
3. Open a **new** terminal and sign in to each one once:
   - `claude`: pick your Claude Pro account, then `/exit`
   - `agy`: sign in with your Google AI Pro account
   - `copilot`: run `/login`, sign in with GitHub, then `/exit`
4. Double-click **run_council.bat**.

## Microsoft 365 Copilot setup (once, about 10 minutes)

The Copilot API only accepts apps you've registered in your business account.

1. Go to **https://entra.microsoft.com** and sign in with your **business** account (*you@yourname.onmicrosoft.com*).
2. In the left menu, open **Entra ID > App registrations** and click **+ New registration**.
   - Name: `Council`
   - Supported account types: **Accounts in this organizational directory only**
   - Leave Redirect URI empty, then click **Register**.
3. On the app's **Overview** page, copy the **Application (client) ID** and the **Directory (tenant) ID**. Council asks for both the first time you use the M365 pane.
4. Open **Authentication**. Under *Advanced settings*, set **Allow public client flows** to **Yes**, then click **Save**.
5. Open **API permissions > + Add a permission > Microsoft Graph > Delegated permissions** and add all 7 of these. The API refuses to work if any is missing.
   `Sites.Read.All`, `Mail.Read`, `People.Read.All`, `OnlineMeetingTranscript.Read.All`, `Chat.Read`, `ChannelMessage.Read.All`, `ExternalItem.Read.All`
6. Still on API permissions, click **Grant admin consent for (your org)**, then **Yes**. Every row should show a green check.
7. Check that your user has the license: **https://admin.microsoft.com > Users > Active users >** you **> Licenses and apps**. Both *Microsoft 365 Business Standard* and *Microsoft 365 Copilot* should be ticked.

The first time you ask a question with the M365 pane on, Council opens a Microsoft sign-in page and copies a short code for you. Paste the code there, sign in with your business account, and you stay signed in after that.

## Using it

- Type a question and press **Enter** (Shift+Enter adds a new line).
- Untick a model to leave it out.
- **Combine with** sends all the answers to the model you pick, which writes one merged answer.
- Follow-up questions remember the last 6 exchanges for each model.
- **New chat** starts fresh. Every chat is saved as Markdown in `%USERPROFILE%\.council\chats`.

## Good to know

- Your usage counts against each plan's normal limits. If a pane shows a limit or sign-in error, that plan is the one to look at.
- If `ANTHROPIC_API_KEY` is set in Windows, Claude Code would bill that key instead of using your Pro plan. Council removes it before it runs Claude, so it always uses your subscription.
- The tools run inside an empty folder (`%USERPROFILE%\.council\workspace`) and are told to only chat, so they don't touch your files.
- **Settings** opens `config.json`. There you can change the timeout, how much history each model keeps, or a model's command (for example to pick a specific model with `--model`).
- M365 Copilot keeps its own conversation, so follow-ups work, and **New chat** starts a fresh one. Its API is still in preview, so Microsoft may change it.
- **Settings** can also sign you out of Microsoft 365. Do this if the M365 pane shows a 401 error, or after you change permissions.
- If the Gemini pane says it "printed nothing", run `agy update`. Older agy versions had a bug that dropped output when run from another program.
