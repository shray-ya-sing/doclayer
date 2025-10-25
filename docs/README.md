# DocLayer Documentation

This directory contains the VitePress documentation site for DocLayer.

## Local Development

```bash
# Install dependencies
npm install

# Start dev server
npm run docs:dev
```

The documentation will be available at `http://localhost:5173`

## Building

```bash
# Build for production
npm run docs:build

# Preview production build
npm run docs:preview
```

## Deploying to Vercel

### Step 1: Push to GitHub

Make sure this documentation is committed and pushed to your repository.

### Step 2: Create New Vercel Project

1. Go to [Vercel Dashboard](https://vercel.com/dashboard)
2. Click "Add New" → "Project"
3. Import your `doclayer` repository

### Step 3: Configure Project Settings

**IMPORTANT:** Configure these settings:

- **Framework Preset**: VitePress (auto-detected)
- **Root Directory**: `docs` ← Set this!
- **Build Command**: `npm run docs:build` (auto-filled)
- **Output Directory**: `.vitepress/dist` (auto-filled)
- **Install Command**: `npm install` (auto-filled)

### Step 4: Deploy

Click "Deploy" and wait for build to complete (~1-2 minutes).

### Step 5: Configure Custom Domain

1. Go to Project Settings → Domains
2. Add your custom domain (e.g., `docs.yourdomain.com`)
3. Follow Vercel's DNS instructions

**For subdomain:**
- Add CNAME record: `docs.yourdomain.com` → `cname.vercel-dns.com`

### Step 6: Enable Auto-Deploy

Vercel will automatically redeploy when you push to your main branch.

## Project Structure

```
docs/
├── .vitepress/
│   └── config.mts         # VitePress configuration
├── guide/
│   ├── introduction.md
│   ├── installation.md
│   └── getting-started.md
├── api/
│   ├── csharp.md
│   ├── python.md
│   ├── typescript.md
│   └── webapi.md
├── index.md             # Homepage
├── package.json
├── vercel.json
└── README.md            # This file
```

## Troubleshooting

### Build fails on Vercel

1. Check that **Root Directory** is set to `docs`
2. Verify Node.js version (should be 18.x or higher)
3. Check build logs in Vercel dashboard

### Pages not loading

1. Verify all markdown files have correct frontmatter
2. Check that links in sidebar match file paths
3. Ensure all referenced files exist

### Search not working

Search is enabled by default with VitePress local search. No additional configuration needed.

## Adding New Pages

1. Create a new `.md` file in the appropriate directory
2. Add the page to sidebar in `.vitepress/config.mts`
3. Commit and push - Vercel will auto-deploy

Example:

```typescript
// .vitepress/config.mts
sidebar: [
  {
    text: 'Your Section',
    items: [
      { text: 'New Page', link: '/path/to/new-page' }
    ]
  }
]
```

## Features

- Fast build times (<2 minutes)
- Auto-deployment on git push
- Local search included
- Mobile responsive
- Dark mode support
- Syntax highlighting
- Code group tabs
- GitHub integration

## Support

For issues with documentation:
- Create an issue in the main repository
- Check VitePress docs: https://vitepress.dev
