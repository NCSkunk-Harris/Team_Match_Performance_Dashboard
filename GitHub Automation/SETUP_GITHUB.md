# One-time GitHub setup

Your project is already a git repo with an initial commit. Follow these steps
once to push it to GitHub and turn on the live page + auto-rebuild.

## 1. Create an empty repo on GitHub

Go to https://github.com/new and:
- **Repository name:** `courage-performance-dashboard` (or your choice)
- **Visibility:** Private (recommended) or Public
- **Do NOT** check "Add a README", ".gitignore", or "license" — the repo
  already has these. Leave it empty.
- Click **Create repository**.

## 2. Connect your local folder and push

Open Terminal, then run these (replace `YOUR-USERNAME` and the repo name if
you changed it):

```bash
cd "/Users/tomharris/Desktop/Claude/Projects/Team Performance Dashboard"
git remote add origin https://github.com/YOUR-USERNAME/courage-performance-dashboard.git
git push -u origin main
```

If prompted to log in, use your GitHub username and a **Personal Access Token**
as the password (GitHub no longer accepts your account password here).
Create one at: https://github.com/settings/tokens → "Generate new token (classic)"
→ check the `repo` scope → copy it and paste when asked.

## 3. Enable GitHub Pages (the live URL)

In your repo on GitHub:
1. Go to **Settings → Pages**.
2. Under **Build and deployment → Source**, select **GitHub Actions**.
3. Save.

The included workflow will run on your next push, rebuild the dashboard, and
publish it. Your live URL will be:

```
https://YOUR-USERNAME.github.io/courage-performance-dashboard/
```

(It also appears under the **Actions** tab → latest run → deploy step.)

## 4. Done — your new weekly workflow

After the one-time setup, each week:

```bash
cd "/Users/tomharris/Desktop/Claude/Projects/Team Performance Dashboard"
# (update the spreadsheet / MANUAL_OVERRIDES, then:)
git add -A
git commit -m "Add match week 12"
git push
```

GitHub rebuilds and republishes the dashboard automatically. To bring in a
teammate, share the repo (**Settings → Collaborators**) and have them follow
`CONTRIBUTING.md`.
```
```
