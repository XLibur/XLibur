# XLibur documentation site

The XLibur documentation site, built with [Docusaurus](https://docusaurus.io/).
Published to GitHub Pages at <https://xlibur.github.io/XLibur/>.

## Local development

This project uses [pnpm](https://pnpm.io/). The version is pinned by the
`packageManager` field in [`package.json`](package.json), so `corepack enable` is
enough to get the right one.

```sh
pnpm install
pnpm start
```

`pnpm start` serves the site at <http://localhost:3000/XLibur/> with hot reload.

## Build

```sh
pnpm run build      # static output in ./build
pnpm run serve      # serve the built output locally
pnpm run typecheck  # type check the config/sidebars
```

## Content

Documentation pages live in [`docs/`](docs) as Markdown. Ordering is controlled by
[`sidebars.ts`](sidebars.ts); each page's `sidebar_position` front matter is a fallback.

The site runs in **docs-only mode** — `docs/introduction.md` has `slug: /` and is served
as the site root, so there is no separate landing page under `src/pages`.

## Publishing

[`.github/workflows/docs.yml`](../.github/workflows/docs.yml) builds the site on every push
to `main` that touches `docs-website/`, and deploys it to GitHub Pages. Pull requests build
the site (and type check) without deploying.

One-time repository setup: **Settings → Pages → Build and deployment → Source: GitHub Actions**.
