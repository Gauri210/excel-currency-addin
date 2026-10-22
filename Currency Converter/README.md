
# Currency Converter - Excel Add-in

IMAP Group Project, Topic: *Excel Add-in that Fetches API Data*.

An Office task pane add-in for Excel. The user enters an amount and picks two currencies, and the add-in fetches the exchange rate from the Frankfurter API, calculates the converted amount, and writes the result with the rate and its date into the worksheet.

**Tech:** HTML, CSS and JavaScript task pane, Office.js, Frankfurter API (no API key), Node.js, webpack, Yeoman Office generator.

---

## Member 1: Setup, Manifest and Sideloading (Gauri Deshmukh)

**Branches:** `feat/setup-manifest`, `docs/readme-setup`
**Files:** `manifest.xml`, `package.json`, `webpack.config.js`, `.gitignore`, `README.md`

### What I did

1. Created the GitHub repo, added the team as collaborators, and protected `main` with a ruleset (pull request and one approval required, no force pushes or deletion).
2. Generated the add-in with the Yeoman Office generator:

```bash
   npm install -g yo generator-office
   yo office
```

   Choices: Office Add-in Task Pane project, JavaScript, add-in only (XML) manifest, name "Currency Converter", Excel.
3. Added a `.gitignore` (`node_modules/`, `venv/`, `.env`, `*.pem`, `*.crt`).
4. Customised `manifest.xml`: display name, description, provider name, ribbon group ("Currency"), button label ("Convert Currency"), tooltip and get-started text. The add-in `Id` and the localhost addresses were left unchanged.
5. Checked the manifest with `npm run validate` and ran the add-in with `npm start`.
6. Pushed the scaffold and manifest through a pull request, and shared the sideloading steps with the team.

### How to run and sideload

**Desktop Excel**

```bash
npm install
npm start
```

This builds the add-in, starts the dev server on `https://localhost:3000`, and sideloads the add-in into Excel. Accept the certificate prompt if it appears. Stop it with `npm stop`.

**Excel on the web (fallback)**

1. Keep `npm start` running so the dev server is up.
2. In Excel on the web: Insert, Add-ins, Upload My Add-in.
3. Choose `manifest.xml` from the project folder.

**Troubleshooting**

- Certificate error: run `npx office-addin-dev-certs install`, then restart Excel.
- Add-in not showing: close all Excel windows and run `npm start` again.
- Old version still showing: run `npm stop`, then `npm start`.

### Research answers

#### What are Office Add-ins and the manifest file?

_Write your answer here in your own words. Cover: what an Office Add-in is (a web app running inside an Office app, using Office.js); what the manifest is; its key parts (`Id`, `DisplayName`, `Hosts`, `SourceLocation`, `Permissions`, `VersionOverrides` for the ribbon button); and why we chose the add-in only XML manifest._

#### How do you sideload an add-in for testing?

_Write your answer here in your own words. Cover: what sideloading means; how `npm start` does it on desktop; how to upload the manifest on Excel on the web; and why HTTPS and a certificate are needed._

### Pull requests

- `feat/setup-manifest` to `main`: (add link)
- `docs/readme-setup` to `main`: (add link)
