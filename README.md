# Email Exporter for New Outlook

A small Outlook web add-in that exports metadata from messages selected in **New Outlook for Windows** or **Outlook on the web** directly to an **Excel `.xlsx` workbook**.

Current version: **1.3.0**

## What it exports

Default fields:

1. From Name
2. From Email
3. To Name
4. To Email
5. Subject

Optional fields include Sender / Sent-by details, To Source, Reply-To, Return-Path, Internet Message ID, and Date Header.

For spam analysis, the add-in treats the MIME `From:` and `To:` headers as the authoritative source for From/To names and email addresses. Outlook structured address fields are used only as fallbacks.

## Privacy

This project has **no backend service**. The hosted site contains only static HTML, CSS, JavaScript, and icons. Selected message metadata is read inside the Outlook task pane and the XLSX workbook is generated locally in the browser. The add-in does not POST mailbox data to this GitHub Pages site.

Outlook requires the manifest permission `ReadWriteMailbox` for `loadItemByIdAsync`. This add-in only reads message data and does not modify messages.

## Outlook requirements and limits

- New Outlook for Windows or Outlook on the web
- Mailbox API **1.15**
- Up to **100 selected messages per export**
- HTTPS hosting
- Custom add-in sideloading enabled for the Outlook account

## Repository layout

```text
outlook-email-exporter/
├── .github/workflows/pages.yml   # Publishes /web to GitHub Pages
├── web/                          # Production static site
│   ├── taskpane.html
│   ├── taskpane.js
│   ├── taskpane.css
│   ├── xlsx-lite.js
│   └── assets/
├── manifest.xml                  # Production manifest: GitHub Pages
├── manifest.local.xml            # Localhost development manifest
├── manifest.template.xml         # Template for another HTTPS host
├── server.js                     # Local HTTPS dev server
├── package.json
├── configure-host.ps1
├── CHANGELOG.md
└── README.md
```

# GitHub Pages deployment

This repository is prepared for:

- GitHub user: `AndreLees`
- Repository: `outlook-email-exporter`
- Hosted URL: `https://andrelees.github.io/outlook-email-exporter/`
- Task pane: `https://andrelees.github.io/outlook-email-exporter/taskpane.html`

The production `manifest.xml` already points to that URL.

## 1. Create the repository

Create a new GitHub repository named exactly:

```text
outlook-email-exporter
```

For GitHub Free, make the repository **Public** if you want to use GitHub Pages without upgrading.

Do not initialise it with a README, `.gitignore`, or licence if you are uploading this prepared project as the initial contents.

## 2. Upload/push this project

Upload all files and folders in this repository package to the repository's `main` branch.

The important hidden folder is:

```text
.github/workflows/pages.yml
```

It must be included because it performs the Pages deployment.

## 3. Enable GitHub Pages

In GitHub:

**Repository → Settings → Pages → Build and deployment → Source → GitHub Actions**

The included workflow publishes only the `web` directory.

After the workflow succeeds, the add-in files will be available at:

```text
https://andrelees.github.io/outlook-email-exporter/
```

## 4. Install the production manifest in Outlook

Sideload the repository's:

```text
manifest.xml
```

In New Outlook:

**More Apps → Get add-ins → My add-ins → Add a Custom Add-in → Add from File**

After installing the production manifest, the local Node server is no longer required.

## Updating the add-in later

For normal HTML/CSS/JavaScript changes:

1. Change files under `web/`.
2. Commit/push to `main`.
3. GitHub Actions automatically redeploys GitHub Pages.
4. Close and reopen the Outlook task pane if Outlook has cached an older copy.

You normally **do not need to reinstall `manifest.xml`** unless the manifest itself changes—for example permissions, version, icons, command configuration, or the hosting URL.

# Local development

The production site and local development are deliberately separated.

## First-time local setup

Install a current Node.js LTS release, then from the repository folder run:

```powershell
npm install
npm run trust-cert
npm run dev
```

The local task pane runs at:

```text
https://localhost:3000/taskpane.html
```

Sideload:

```text
manifest.local.xml
```

This lets you test changes locally before pushing them to GitHub.

## Subsequent local runs

```powershell
npm run dev
```

You normally only need to trust the Office development certificate once per PC.

# Using another HTTPS host

Generate a separate custom manifest without overwriting the committed production manifest:

```powershell
.\configure-host.ps1 -BaseUrl "https://your-host.example/path"
```

This creates:

```text
manifest.custom.xml
```

# Version history

See [CHANGELOG.md](CHANGELOG.md).
