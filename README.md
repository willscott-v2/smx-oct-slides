# SMX GEO Master Class 2025 - Presentation Generator

## Current Project Files

### Active Files (Production Ready)
- **`slides.csv`** - Main slide data (61 slides, pipe-delimited)
- **`generator.js`** - Google Apps Script for slide generation
- **`config.csv`** - Configuration settings (fonts, colors, themes)
- **`promopage.md`** - Promotional page content

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
- 61 comprehensive slides on Generative Engine Optimization
- 4 themed sessions with distinct backgrounds
- Detailed speaker notes (150-200 words per slide)
- Case studies and tool recommendations
- Future trends and measurement strategies

## Last Updated
October 3, 2024

## Contact
Will Scott - Search Influence
willscott.me/links