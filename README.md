# Eunice Delali Alomatu — Portfolio Site

This folder is a ready-to-publish website: `index.html` (the whole site, styling included),
`resume.pdf` (linked from the "Resume" buttons on the page), and `headshot.jpg` (your
photo in the header and hero section). No build step, no dependencies — GitHub Pages
will serve it exactly as-is.

You already have a live preview link in the chat. These steps put the same site at a
permanent URL you control, e.g. `https://<your-username>.github.io`.

## 1. Create the GitHub repository

1. Go to [github.com/new](https://github.com/new).
2. **Repository name** — you have two options:
   - `<your-username>.github.io` (replace `<your-username>` with your actual GitHub
     username, exactly). This is the simplest option: the site becomes your root
     profile site, live at `https://<your-username>.github.io` with no extra setup.
   - Any other name (e.g. `portfolio`). The site will instead live at
     `https://<your-username>.github.io/portfolio/`.
3. Leave it **Public**, don't initialize it with a README (you already have one here),
   and click **Create repository**.

## 2. Push these files

Open a terminal, `cd` into this folder, and run:

```bash
git init
git add index.html resume.pdf headshot.jpg
git commit -m "Add portfolio site"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo-name>.git
git push -u origin main
```

Replace `<your-username>` and `<your-repo-name>` with the values from Step 1. GitHub
will prompt you to sign in (or use a personal access token) the first time you push.

## 3. Turn on GitHub Pages

1. On GitHub, open your new repository's **Settings** tab.
2. In the left sidebar, click **Pages**.
3. Under **Build and deployment → Source**, choose **Deploy from a branch**.
4. Under **Branch**, choose `main` and folder `/ (root)`, then **Save**.
5. Wait about a minute, then refresh the page — GitHub will show the live URL at the
   top (either `https://<your-username>.github.io` or
   `https://<your-username>.github.io/<your-repo-name>/`, depending on what you named
   the repo in Step 1).

That URL is what you put on your resume, LinkedIn, or job applications.

## Making changes later

Edit `index.html` directly (it's plain HTML/CSS/JS, no build tools needed), then:

```bash
git add index.html
git commit -m "Update portfolio"
git push
```

GitHub Pages redeploys automatically within a minute or two of every push.

## What's on the site

- **Profile** — a short summary of your background
- **Experience** — Amazon, First National Bank, Access Bank, Agricultural Development Bank
- **Education** — Haas MBA and your undergraduate degree
- **Skills & Certifications** — your chartered credentials and technical skills
- **Leadership & Community**
- **Contact** — email (click to copy), LinkedIn, and a link to download `resume.pdf`
- **Photo** — your headshot (`headshot.jpg`) in the header and hero section

Everything is sourced directly from your resume. If anything needs correcting or you'd
like a different section added (projects, publications), just ask and the `index.html`
file can be updated.
