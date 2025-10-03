# GitHub Repository Setup for SMX GEO Generator (With Versioning)

## Steps to Create Repository:

1. **Create GitHub Repository:**
   ```bash
   # Go to GitHub.com and create new repository
   Repository name: smx-geo-generator
   Description: SMX GEO Master Class Presentation Generator with Auto-Updates
   Public/Private: Your choice
   Initialize with README: Yes
   Add .gitignore: None needed
   Choose license: MIT (recommended)
   ```

2. **Clone and Setup Locally:**
   ```bash
   git clone https://github.com/YOUR_USERNAME/smx-geo-generator.git
   cd smx-geo-generator
   ```

3. **Add Your Files with Versioning:**
   ```bash
   # Copy core files
   cp "/Users/willscott/Cursor-nd/SMX Master Class 202510/config.csv" .
   cp "/Users/willscott/Cursor-nd/SMX Master Class 202510/slides.csv" .
   cp "/Users/willscott/Cursor-nd/SMX Master Class 202510/generator.js" .
   cp "/Users/willscott/Cursor-nd/SMX Master Class 202510/README.md" .
   cp "/Users/willscott/Cursor-nd/SMX Master Class 202510/version.json" .
   ```

4. **Create Version Tags and Push:**
   ```bash
   git add .
   git commit -m "feat: Initial release v1.0.0 - SMX GEO Generator with auto-updates"
   git tag v1.0.0
   git push origin main
   git push origin v1.0.0
   ```

## Versioning Strategy:

### Semantic Versioning (SemVer):
- **Major (X.0.0)**: Breaking changes, major new features
- **Minor (1.X.0)**: New features, backwards compatible
- **Patch (1.0.X)**: Bug fixes, small improvements

### GitHub Branches:
- **main**: Latest stable version
- **develop**: Development branch for testing
- **v1.x**: Stable release branches

### Example Version Updates:
```bash
# For new slides or major content changes
git tag v1.1.0
git push origin v1.1.0

# For bug fixes or small tweaks
git tag v1.0.1
git push origin v1.0.1
```

## Configuration Update (Required):

Update these settings in your generator.js:

```javascript
const GITHUB_CONFIG = {
  username: 'YOUR_ACTUAL_USERNAME',    // ← Change this
  repo: 'smx-geo-generator',
  branch: 'main',
  currentVersion: '1.0.0'
};
```

## Testing Plan:

### Phase 1 - Safe Testing (Current)
- ✅ Preview function shows simulated changes
- ✅ Apply function creates backups but doesn't update
- ✅ Menu items added for testing

### Phase 2 - Live Testing
1. Create GitHub repository
2. Update URLs in code
3. Test Preview function with real GitHub data
4. Test Apply function (with backups)

### Phase 3 - Production
- Remove simulation mode
- Enable full auto-update functionality

## File Structure in GitHub:
```
smx-geo-generator/
├── README.md          # Documentation
├── config.csv         # Configuration settings
├── slides.csv         # Slide data (61 slides)
├── generator.js       # Google Apps Script
└── docs/              # Additional documentation
    └── setup-guide.md # Setup instructions
```