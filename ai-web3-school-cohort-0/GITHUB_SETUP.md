# GitHub Setup

The Learning Agent startup prompt requires confirmation before creating or publishing a repository. Use this file as the safe checklist.

## Confirm First

- GitHub username:
- Repository name: `ai-web3-school-cohort-0`
- Visibility: public recommended for proof-of-work
- Default branch: `main`
- License: optional

## Option A: GitHub Website

1. Create a new repository on https://github.com/new.
2. Name it `ai-web3-school-cohort-0`.
3. Choose public visibility only after reviewing all files for secrets.
4. Upload or push this directory.

## Option B: GitHub CLI

Install Git and GitHub CLI first:

- Git: https://git-scm.com/downloads
- GitHub CLI: https://cli.github.com/

Then run from this directory:

```bash
git init
git branch -M main
git add .
git commit -m "Initialize AI Web3 School learning repo"
gh auth login
gh repo create ai-web3-school-cohort-0 --public --source=. --remote=origin --push
```

## Safety Check Before Push

```bash
git status
git diff --cached
```

Review the diff and confirm:

- No API keys.
- No private keys or seed phrases.
- No private contact details.
- No internal links.
- No unreleased business information.

## First Issue Ideas

- Confirm learner profile and select first-week path.
- Create first Handbook feedback issue.
- Track first WCB check-in.
