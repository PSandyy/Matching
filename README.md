# Pharmacy Master Matcher

AI-powered pharmacy duplicate detection with multiple file upload support and Google Maps verification.

## Features

- 📁 **Multiple file upload** - CSV/Excel files supported
- 🤖 **AI matching** - Uses Claude AI to identify duplicate pharmacies
- 🗺️ **Google Maps integration** - Verify addresses with one click
- 🏷️ **Smart grouping** - Duplicate pharmacies share the same master code
- 📊 **CSV export** - Download matched results

## Quick Start

1. Install dependencies:
```bash
npm install
```

2. Run locally:
```bash
npm start
```

## API Key

This app calls the Anthropic API directly from the browser. On first use, enter
your Anthropic API key in the sidebar ("🔑 Anthropic API Key"). It is stored
only in your browser's localStorage and is never sent anywhere except
`api.anthropic.com`.

## Deploying to GitHub Pages

1. Make sure `homepage` in `package.json` matches your GitHub Pages URL
   (e.g. `https://<username>.github.io/<repo-name>`).
2. Deploy:
```bash
npm run deploy
```
This builds the app and pushes the `build` folder to the `gh-pages` branch.

3. In your repo settings, set GitHub Pages source to the `gh-pages` branch.
