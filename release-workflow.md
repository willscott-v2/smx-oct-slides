# Release Workflow for SMX GEO Generator

## Automated Version Management

### When to Create New Versions:

**Patch (1.0.X)** - Bug fixes, typos, small improvements:
- Fix speaker notes
- Correct slide formatting
- Minor script improvements

**Minor (1.X.0)** - New features, content additions:
- Add new slides
- New functionality in generator
- Additional configuration options

**Major (X.0.0)** - Breaking changes:
- Complete slide restructure
- Major API changes
- New file format requirements

### Release Checklist:

1. **Test Changes Locally:**
   ```bash
   # Test your changes
   # Verify all functions work
   # Check slide generation
   ```

2. **Update version.json:**
   ```json
   {
     "version": "1.1.0",
     "releaseDate": "2024-10-03",
     "changes": [
       "Added 5 new slides on AI tools",
       "Improved error handling",
       "Fixed formatting issues"
     ]
   }
   ```

3. **Commit and Tag:**
   ```bash
   git add .
   git commit -m "feat: Release v1.1.0 - Enhanced content and functionality"
   git tag v1.1.0
   git push origin main
   git push origin v1.1.0
   ```

4. **Create GitHub Release:**
   - Go to GitHub → Releases → New Release
   - Choose tag: v1.1.0
   - Title: "SMX GEO Generator v1.1.0"
   - Description: Copy from version.json changes
   - Attach files if needed

### Auto-Update Testing Process:

1. **Test Preview Function:**
   - Run "Preview Updates (Safe)" in Google Sheets
   - Verify version detection works
   - Check change summary

2. **Test Apply Function:**
   - Run "Apply Updates"
   - Verify backup creation
   - Confirm data updates correctly

3. **Rollback if Needed:**
   - Use backup sheets created automatically
   - Revert to previous GitHub tag/branch

### Distribution Strategy:

**For End Users:**
- Use `main` branch (always stable)
- Auto-update pulls from latest tag
- Clear release notes for each version

**For Testing:**
- Use `develop` branch
- Test new features before release
- Preview mode shows what would change

**For Enterprise:**
- Use specific version tags (e.g., v1.0.0)
- Controlled update schedule
- Manual approval process

### Version.json Template:

```json
{
  "version": "1.X.X",
  "releaseDate": "YYYY-MM-DD",
  "description": "Brief description of this release",
  "changes": [
    "List of changes in this version",
    "One bullet per major change",
    "Focus on user-facing improvements"
  ],
  "compatibility": {
    "minGoogleSheetsVersion": "current",
    "minAppsScriptVersion": "current"
  },
  "breaking_changes": [
    "List any breaking changes here",
    "Include migration instructions"
  ],
  "migration_notes": "Instructions for upgrading from previous version"
}
```

### GitHub Release Automation (Optional):

Create `.github/workflows/release.yml` for automated releases:

```yaml
name: Create Release
on:
  push:
    tags:
      - 'v*'
jobs:
  release:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v2
      - name: Create Release
        uses: actions/create-release@v1
        with:
          tag_name: ${{ github.ref }}
          release_name: SMX GEO Generator ${{ github.ref }}
          body_path: version.json
```