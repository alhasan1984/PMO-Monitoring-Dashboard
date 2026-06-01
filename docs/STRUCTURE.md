# PMO Monitoring Dashboard Structure

This repository is intentionally static so it can be hosted with GitHub Pages or copied to an internal web server.

## Folders

- `src/`: Dashboard application files.
- `src/assets/`: Oracle theme assets, customer logos, and mapped people images.
- `scripts/`: Local helper scripts for serving the app during development.
- `.github/workflows/`: GitHub Pages deployment workflow.

## Deployment

The GitHub Pages workflow publishes the `src` folder. Any file needed by the browser must live inside `src`.
