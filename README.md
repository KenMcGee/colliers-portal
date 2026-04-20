# Colliers Denver Design Studio

A professional CRE document portal for generating Offering Memoranda and Broker Opinions of Value.

## Features
- OM / BOV Builder with 5-step guided workflow
- AI-powered narrative generation (requires Claude API key)
- Financial calculator (NOI, cap rate, price/SF, GRM)
- Rent roll entry
- File upload support
- Saved projects dashboard
- Print-to-PDF document output

## Deployment to Vercel (Free Tier)

### Option 1 — Vercel CLI (recommended)
1. Install Node.js from https://nodejs.org
2. Open Terminal and run: `npm install -g vercel`
3. Navigate to this folder: `cd colliers-portal`
4. Run: `vercel`
5. Follow the prompts — choose "No" for existing project, accept defaults
6. Your app will be live at a `.vercel.app` URL

### Option 2 — Vercel Dashboard (no CLI needed)
1. Go to https://vercel.com and sign up for a free account
2. Click "Add New Project"
3. Choose "Upload" and drag this entire folder onto the upload area
4. Click Deploy
5. Done — live in ~60 seconds

### Option 3 — GitHub + Vercel (best for ongoing updates)
1. Create a free GitHub account at https://github.com
2. Create a new repository called `colliers-portal`
3. Upload all files in this folder to the repository
4. Go to https://vercel.com, click "Add New Project"
5. Import your GitHub repository
6. Click Deploy
7. Future updates: edit files in GitHub → Vercel auto-redeploys

## Setting Up the Claude API Key
1. Go to https://console.anthropic.com
2. Create an account and generate an API key
3. In the portal, go to Settings and paste your key
4. The key is stored locally in the browser — never transmitted except to Anthropic

## Adding Tools
As new tools are built (Map Generator, Brochure Builder, etc.), 
add their HTML pages and link them in the sidebar nav in index.html.

## Customization
- Colors: Edit CSS variables at the top of styles.css
- Firm name/disclaimer: Update via the Settings page in the portal
- Logo: Replace the "C" in .logo-mark with an <img> tag

## Files
- index.html — Main portal shell and all page HTML
- styles.css — All styling and Colliers brand colors
- app.js — All application logic, API calls, document generation
- vercel.json — Vercel deployment configuration
