# How to Push This Repo to GitHub

Follow these steps to get your Toastmasters project live on GitHub. You only need to do this once.

---

## Step 1: Create the Repository on GitHub

1. Go to **https://github.com/new** (log in if needed)
2. Fill in:
   - **Repository name**: `sierra-speakers-toastmasters`
   - **Description**: `Google Apps Script macros for Toastmasters meeting management — role confirmations, agenda generation, and AI-powered club emails`
   - **Public** (so others can see and learn from it)
   - **Do NOT** check "Add a README" or ".gitignore" or "License" — we already have those locally
3. Click **Create repository**
4. You'll see a page with setup instructions — leave this tab open, you'll need the URL

---

## Step 2: Open a Terminal in This Folder

Open **PowerShell** or **Command Prompt** and navigate to this folder:

```powershell
cd "C:\Users\rakow\OneDrive\Desktop\Projects\Toastmasters\github-repo"
```

If you don't have Git installed, download it from https://git-scm.com/downloads and install with the default settings. Then close and reopen your terminal.

---

## Step 3: Initialize Git and Make Your First Commit

Run these commands one at a time:

```bash
git init
```

```bash
git add .
```

```bash
git commit -m "Initial commit: Sierra Speakers Toastmasters meeting management macros"
```

---

## Step 4: Connect to GitHub and Push

```bash
git branch -M main
```

```bash
git remote add origin https://github.com/rakowski7/sierra-speakers-toastmasters.git
```

```bash
git push -u origin main
```

If this is your first time using Git with GitHub, it will ask you to authenticate. A browser window should open — just follow the prompts to log in.

---

## Step 5: Verify

Go to **https://github.com/rakowski7/sierra-speakers-toastmasters** — you should see all your files there with the README displayed on the main page.

---

## Step 6 (Optional): Protect the Main Branch

This prevents accidental pushes directly to `main` — all changes would need to go through a Pull Request.

1. Go to your repository on GitHub
2. Click **Settings** (tab at the top)
3. In the left sidebar, click **Branches** (under "Code and automation")
4. Click **Add branch protection rule** (or **Add rule**)
5. Fill in:
   - **Branch name pattern**: `main`
   - Check **Require a pull request before merging**
   - Optionally check **Require approvals** (set to 1 if you want a review step, or leave unchecked if you're the only contributor)
6. Click **Create** or **Save changes**

Now direct pushes to `main` will be blocked, and you'll create feature branches and merge via Pull Requests instead.

---

## Day-to-Day Workflow (After Setup)

When you make changes to the code:

```bash
# Check what changed
git status

# Stage your changes
git add .

# Commit with a message describing what you changed
git commit -m "Add support for virtual-only meetings"

# Push to GitHub
git push
```

That's it! Your code is version-controlled and backed up on GitHub.

---

## Connecting clasp (Optional — Advanced)

If you want to push/pull code between this repo and the actual Apps Script project:

1. Install Node.js from https://nodejs.org/
2. Install clasp:
   ```bash
   npm install -g @google/clasp
   ```
3. Log in:
   ```bash
   clasp login
   ```
4. Find your Script ID:
   - Open the Apps Script editor (`Extensions > Apps Script` in your spreadsheet)
   - Go to **Project Settings** (gear icon on the left)
   - Copy the **Script ID**
5. Edit `.clasp.json` in this folder and replace `YOUR_SCRIPT_ID_HERE` with your actual Script ID
6. Now you can:
   ```bash
   clasp push   # Upload local code to Apps Script
   clasp pull   # Download code from Apps Script to local
   ```
