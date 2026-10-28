# PharmEase — Pharmacy POS & Inventory

Installable web app (works online and offline). Data lives in Firebase (Firestore + sign-in).
The app files are hosted with Firebase Hosting.

## Folder map

```
pharmease/
├── README.md                      this file
├── firebase.json                  deployment settings (keep in the main folder)
├── pwa/                           the app itself — this folder is what gets published
│   ├── index.html                 the whole app (put your Firebase settings here, see Step 1)
│   ├── sw.js                      makes the app open offline (change the version number when you update)
│   ├── manifest.webmanifest       app name, colors and icon for installing
│   └── icons/                     app icons (192, 512, maskable)
├── firebase/                      database settings, deployed with the command below
│   ├── firestore.rules            who can read and write what (security rules)
│   ├── firestore.indexes.json     the index for sales lists
│   └── firestore-rules.txt        plain-text copy of the rules, for pasting by hand
└── old-apps-script/               the first (Google Sheets) version, kept for reference only
```

## First-time setup

1. **Firebase settings.** Open `pwa/index.html`, search for `PASTE_API_KEY`, and replace each
   `PASTE_...` value with your project's settings (Firebase console > Project settings > Your apps > Config).
   Keep the quotes and commas. Save.
2. **Install the tools** (once): Node.js (LTS), then in a terminal: `npm install -g firebase-tools` and `firebase login`.
3. **Open a terminal in this folder** (the one containing `firebase.json`) and choose your project (once):
   `firebase use --add`
4. **Deploy:**
   ```
   firebase deploy --only firestore:rules,firestore:indexes,hosting
   ```
   If it asks to delete an old index, answer **N**.
5. Open the address it prints (`https://<project-id>.web.app`). The first screen asks you to create the **Owner** account
   (PIN of 6 to 8 digits). Do this right away.
6. In Firebase console > Authentication > Settings > Authorized domains, make sure your `web.app` address is listed.

## Updating later

1. Replace the changed files. Paste your Firebase settings into `pwa/index.html` again if you replaced it.
2. In `pwa/sw.js`, raise the number in `const CACHE = 'pharmease-v8'` (v9, v10 ...) so devices pick up the update.
3. `firebase deploy --only hosting` (add `firestore:rules` if the rules file changed).
4. Press Ctrl + Shift + R on the app page.

## Test on your own computer

**Safe sandbox (recommended).** A practice copy of the database runs on your computer. Nothing real is touched,
and the data disappears when you stop it. It needs Java (JDK 17 or newer) installed once.

```
cd /d "C:\Users\BTCSI Library\Desktop\pharmease"
firebase emulators:start --project demo-pharmease
```
Then open  **http://localhost:5000/?emulator**  (the `?emulator` part matters). You will see an orange
"SANDBOX" label. The first screen asks you to create an Owner account (use any 6-digit PIN).
The data viewer is at http://localhost:4000. Stop with Ctrl + C.

**Your real project, locally (no Java).** Publish the rules first so the newest features are allowed:
`firebase deploy --only firestore:rules`, paste your Firebase settings into `pwa/index.html`, then run
`firebase emulators:start --only hosting` and open http://localhost:5000 . Anything you save is real data.

## Good to know

- Cashier accounts see only POS / Checkout and never see stock numbers, costs or markups. Only owners see everything.
- Only owners can change PINs, edit usernames, and delete users (Users & Access).
- Receipts: set auto-print and cash drawer under the POS page > Printer & drawer (owner only, per device).
- Expenses & Deposits (owner only): record expenses by category, record optional daily bank deposits, and see net profit in Reports (pick any month with the month box and the arrows).
- Removing damaged or expired stock (Inventory > Adjust) is recorded automatically as an expense, valued at cost.
- Make regular backups with the Export buttons (products, sales, deliveries, expenses, cashbook).
- This system is not BIR-accredited. Confirm tax and receipt requirements with an accountant.
