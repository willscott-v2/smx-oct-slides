# SMX GEO Master Class 2025 - Presentation Generator

## Current Project Files

### Active Files (Production Ready)
- **`slides.csv`** - Main slide data (87 slides, pipe-delimited)
- **`generator.js`** - Google Apps Script v1.2.0 for slide generation with GitHub auto-update
- **`config.csv`** - Configuration settings (fonts, colors, themes)
- **`promopage.md`** - Promotional page content
- **`chart_prompts.md`** - Detailed chart creation prompts for AI tools

### Supporting Files
- **`prior_slides.md`** - Outline and content structure documentation
- **`version.json`** - Version tracking and metadata
- **`script_version.json`** - Script version information
- **`uma_*.json`** - UMA case study schema files
- **`*.csv` (backup files)** - Various backup versions with timestamps

### How to Use

1. **Setup Google Sheets:**
   - Create a new Google Sheets document
   - Import `slides.csv` as "Slides" tab (use pipe | as delimiter)
   - Import `config.csv` as "Config" tab

2. **Setup Google Apps Script:**
   - In Google Sheets, go to Extensions > Apps Script
   - Replace default code with contents of `generator.js`
   - Save the script

3. **Generate Presentation:**
   - Run `onOpen()` to create menu
   - Use menu: GEO Presentation > Generate Slides
   - Or run `testSlideCreation()` to test single slide first

## Archive Directory

The `archive/` folder contains previous versions and reference materials:
- Previous CSV versions
- Earlier script iterations
- Reference PDFs
- Documentation drafts

## Version Management

When updating files:
1. Copy current version to archive with timestamp: `filename_YYYYMMDD_HHMM.ext`
2. Replace main directory file with new version
3. Update this README if file structure changes

## Key Features
- 87 comprehensive slides on Generative Engine Optimization
- 4 themed sessions with distinct backgrounds
- Detailed speaker notes (150-200 words per slide)
- Case studies and tool recommendations with charts
- Future trends and measurement strategies
- GitHub auto-update integration
- Professional chart generation prompts

## Version Information
- **Current Version:** 1.2.0
- **Last Updated:** October 4, 2025
- **Generator Script:** v1.2.0 with GitHub integration
- **Charts:** 8 detailed chart prompts for AI generation

## Contact
Will Scott - Search Influence
willscott.me/links