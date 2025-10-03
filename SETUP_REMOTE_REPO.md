# Remote Repository Setup - FOLLOW THESE STEPS

## 1. Create GitHub Repository:

**Go to GitHub.com and:**
1. Click the "+" icon → "New repository"
2. **Repository name**: `smx-geo-generator`
3. **Description**: `SMX GEO Master Class Presentation Generator with Auto-Updates`
4. **Visibility**: Choose Public or Private
5. **DO NOT** initialize with README, .gitignore, or license (we already have these)
6. Click "Create repository"

## 2. GitHub will show you a page with setup instructions. IGNORE those and use these instead:

Run these commands in Terminal (from this directory):

```bash
# Add the remote repository (REPLACE YOUR_USERNAME)
git remote add origin https://github.com/YOUR_USERNAME/smx-geo-generator.git

# Set the default branch name
git branch -M main

# Push to GitHub
git push -u origin main

# Push the version tag
git push origin v1.0.0
```

## 3. After pushing, update the generator.js file:

Find this section in generator.js:
```javascript
const GITHUB_CONFIG = {
  username: 'YOUR_USERNAME',        // ← Change to your actual GitHub username
  repo: 'smx-geo-generator',
  branch: 'main',
  currentVersion: '1.0.0'
};
```

Change `YOUR_USERNAME` to your actual GitHub username.

## 4. Test the setup:

1. Go to your GitHub repository URL
2. You should see all files: config.csv, slides.csv, generator.js, etc.
3. Check that the v1.0.0 tag was created (in Releases section)

## 5. Ready to test auto-updates!

Your repository URLs will be:
- Config: `https://raw.githubusercontent.com/YOUR_USERNAME/smx-geo-generator/main/config.csv`
- Slides: `https://raw.githubusercontent.com/YOUR_USERNAME/smx-geo-generator/main/slides.csv`
- Version: `https://raw.githubusercontent.com/YOUR_USERNAME/smx-geo-generator/main/version.json`

## Next Steps:
1. Create a Google Sheets document
2. Import your CSV files
3. Add the generator.js to Apps Script
4. Test the "Preview Updates (Safe)" function!