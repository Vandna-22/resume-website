# Deploying this resume site to GitHub Pages

This repository contains a static resume site. The included GitHub Actions workflow will publish the repository root to GitHub Pages when you push to the `main` branch.

Quick steps to publish:

- Using the `gh` CLI (recommended):

```bash
# create a new repo on GitHub and push the current folder
gh repo create YOUR_USERNAME/REPO_NAME --public --source=. --remote=origin --push
```

- Or manually (replace `YOUR_USERNAME` and `REPO_NAME`):

```bash
git remote add origin git@github.com:YOUR_USERNAME/REPO_NAME.git
git branch -M main
git push -u origin main
```

Notes:
- The workflow file is at `.github/workflows/deploy.yml`.
- If you have a `CNAME` file, GitHub Pages will use that custom domain; ensure DNS is configured.
- After pushing, open the repository on GitHub > Settings > Pages to confirm the site URL and domain configuration.

If you want, I can try to create the GitHub repo and push for you (requires `gh` or network auth). Otherwise run the commands above.
