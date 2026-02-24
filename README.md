# our206 monorepo

This repo is structured as a monorepo. The website lives in:

- `website/`

## Deployment

GitHub Actions deploys `website/` to GitHub Pages via:

- `.github/workflows/deploy-pages.yml`

## Custom domain

The Pages artifact includes:

- `website/CNAME` with `our206.com`

Set your DNS at your domain provider:

1. `A` record for `our206.com` to GitHub Pages IPs:
   - `185.199.108.153`
   - `185.199.109.153`
   - `185.199.110.153`
   - `185.199.111.153`
2. `CNAME` record for `www` to `<your-github-username>.github.io`

Then, in GitHub repo settings:

1. Open `Settings -> Pages`
2. Ensure source is `GitHub Actions`
3. Confirm custom domain is `our206.com`
4. Enable `Enforce HTTPS` after DNS propagates
