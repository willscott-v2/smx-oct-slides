/**
 * Complete GEO Master Class Presentation Generator
 * Version: 1.2.0
 * Last Updated: 2025-10-03
 *
 * CHANGELOG:
 * 1.2.0 (2025-10-03)
 * - Fixed image insertion for all slides with chart_ref or media_ref
 * - Maintains aspect ratio for images and charts
 * - Improved image positioning and sizing
 * - Charts and images now work in any layout type
 *
 * 1.1.0 (2025-10-03)
 * - Added real image insertion from Google Drive
 * - Removed redundant menu functions
 * - Enhanced error handling with retry logic
 * - Fixed image and chart placeholder functions
 *
 * 1.0.0 (2025-10-03)
 * - Initial release with auto-update functionality
 * - GitHub integration for data synchronization
 * - Comprehensive slide generation system
 */

// ============================================================================
// MENU & INITIALIZATION
// ============================================================================

// Script version constant
const SCRIPT_VERSION = '1.2.0';
const SCRIPT_RELEASE_DATE = '2025-10-03';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GEO Presentation')
    .addItem('Generate Slides', 'generatePresentation')
    .addItem('Test Configuration', 'testConfig')
    .addItem('Validate Data', 'validateSlideData')
    .addSeparator()
    .addItem('🔍 Test One Slide', 'testSlideCreation')
    .addSeparator()
    .addItem('🔄 Preview Updates (Safe)', 'previewUpdatesFromGitHub')
    .addItem('📥 Apply Updates', 'applyUpdatesFromGitHub')
    .addSeparator()
    .addItem('ℹ️ Script Version', 'showScriptVersion')
    .addToUi();
}

function showScriptVersion() {
  const ui = SpreadsheetApp.getUi();
  const message =
    `Script Version: ${SCRIPT_VERSION}\n` +
    `Release Date: ${SCRIPT_RELEASE_DATE}\n\n` +
    `Data Version: ${GITHUB_CONFIG.currentVersion}\n\n` +
    `To update the script:\n` +
    `1. Go to Extensions → Apps Script\n` +
    `2. Copy the latest code from GitHub\n` +
    `3. Save and refresh the sheet`;

  ui.alert('Generator Version Info', message, ui.ButtonSet.OK);
}

// ============================================================================
// MAIN GENERATION FUNCTION
// ============================================================================

function generatePresentation() {
  try {
    Logger.log('Starting presentation generation...');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = loadConfig(ss);
    const slides = loadSlides(ss);

    Logger.log(`Loaded: ${slides.length} slides, ${Object.keys(config).length} config items`);

    if (slides.length === 0) {
      throw new Error('No slides found. Check your Slides sheet.');
    }

    // Get the folder of the current spreadsheet
    const spreadsheetFile = DriveApp.getFileById(ss.getId());
    const folder = spreadsheetFile.getParents().next();

    // Create new presentation with configured title
    const presentationTitle = config['deck_title'] || 'GEO Master Class - ' + new Date().toLocaleDateString();
    const presentation = SlidesApp.create(presentationTitle);
    const presentationId = presentation.getId();
    const presentationFile = DriveApp.getFileById(presentationId);
    presentationFile.moveTo(folder);

    Logger.log('Created presentation: ' + presentationTitle);

    // Remove default slide
    const defaultSlides = presentation.getSlides();
    if (defaultSlides.length > 0) {
      defaultSlides[0].remove();
    }

    // Generate slides
    let successCount = 0;
    slides.forEach((slideData, index) => {
      try {
        Logger.log(`Creating slide ${index + 1}: ${slideData.title}`);
        createSlide(presentation, slideData, config);
        successCount++;
      } catch (slideError) {
        Logger.log(`Error creating slide ${index + 1}: ${slideError.toString()}`);
        // Continue with next slide
      }
    });

    // Add footers to all slides except title
    if (config['footer_text']) {
      addFooterToAllSlides(presentation, config);
    }

    SpreadsheetApp.getUi().alert('Success!',
      `Presentation "${presentationTitle}" generated!\n\n${successCount} of ${slides.length} slides created successfully.\n\nOpen: ${presentation.getUrl()}`,
      SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    Logger.log('ERROR: ' + error.toString() + '\n' + error.stack);
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

// ============================================================================
// CONFIGURATION LOADING
// ============================================================================

function loadConfig(ss) {
  const sheet = ss.getSheetByName('Config');
  if (!sheet) {
    throw new Error('Config sheet not found. Please create a Config tab or run Setup Config Tab from the menu.');
  }

  const data = sheet.getDataRange().getValues();
  const config = {};

  // Skip header row, process all config settings
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][1] !== '') {
      // Convert setting name to lowercase with underscores
      const key = data[i][0].toString().toLowerCase().replace(/\s+/g, '_');
      const value = data[i][1];

      // Convert numeric strings to numbers
      if (!isNaN(value) && value !== '') {
        config[key] = Number(value);
      } else if (value === 'true') {
        config[key] = true;
      } else if (value === 'false') {
        config[key] = false;
      } else {
        config[key] = value;
      }
    }
  }

  // Set defaults if not specified
  const defaults = {
    title_slide_font_size: 44,
    section_slide_font_size: 40,
    content_title_font_size: 36,
    subtitle_font_size: 24,
    bullet_font_size: 20,
    header_font: 'Arial',
    body_font: 'Arial',
    text_color: '#000000'
  };

  for (let key in defaults) {
    if (!config[key]) {
      config[key] = defaults[key];
    }
  }

  return config;
}

function loadSlides(ss) {
  const sheet = ss.getSheetByName('Slides');
  if (!sheet) throw new Error('Slides sheet not found');

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('Slides sheet appears to be empty');

  const headers = data[0];
  const slides = [];

  // Create column map - handle both comma and pipe delimited data
  const cols = {};
  headers.forEach((header, index) => {
    const cleanHeader = header.toString().toLowerCase().replace(/\s+/g, '_');
    cols[cleanHeader] = index;
  });

  // Required columns
  const requiredCols = ['order', 'title'];
  for (let reqCol of requiredCols) {
    if (cols[reqCol] === undefined) {
      throw new Error(`Required column '${reqCol}' not found in Slides sheet. Found: ${Object.keys(cols).join(', ')}`);
    }
  }

  // Process each row
  for (let i = 1; i < data.length; i++) {
    if (data[i][cols.order] && data[i][cols.title]) {
      const slide = {
        order: data[i][cols.order],
        section_id: (data[i][cols.section_id] || '').toString(),
        layout: (data[i][cols.layout] || 'Content').toString(),
        title: (data[i][cols.title] || '').toString(),
        subtitle: (data[i][cols.subtitle] || '').toString(),
        bullets: (data[i][cols.bullets] || '').toString(),
        speaker_notes: (data[i][cols.speaker_notes] || '').toString(),
        media_ref: (data[i][cols.media_ref] || '').toString(),
        chart_ref: (data[i][cols.chart_ref] || '').toString()
      };
      slides.push(slide);
    }
  }

  return slides.sort((a, b) => a.order - b.order);
}

// ============================================================================
// SIMPLIFIED SLIDE CREATION
// ============================================================================

function createSlide(presentation, slideData, config) {
  try {
    Logger.log('Creating slide: ' + slideData.title + ' (Layout: ' + slideData.layout + ')');

    // Create slide based on layout, but simplified
    let slide;
    if (slideData.layout === 'Title') {
      slide = createTitleSlide(presentation, slideData, config);
    } else if (slideData.layout === 'Section') {
      slide = createSectionSlide(presentation, slideData, config);
    } else {
      // For all other layouts, use basic title+body
      slide = createBasicSlide(presentation, slideData, config);
    }

    // Add speaker notes
    if (slideData.speaker_notes && slide) {
      try {
        const notesPage = slide.getNotesPage();
        const speakerNotesShape = notesPage.getSpeakerNotesShape();
        speakerNotesShape.getText().setText(slideData.speaker_notes);

        // Style speaker notes
        if (config['speaker_notes_font_size']) {
          speakerNotesShape.getText().getTextStyle()
            .setFontSize(config['speaker_notes_font_size'])
            .setFontFamily(config['body_font'] || 'Arial');
        }
      } catch (notesError) {
        Logger.log('Could not add speaker notes: ' + notesError.toString());
      }
    }

    Logger.log('Slide creation completed successfully');
    return slide;

  } catch (error) {
    Logger.log('Error in createSlide: ' + error.toString());
    throw error;
  }
}

function createTitleSlide(presentation, slideData, config) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE);

  try {
    // Get all shapes and find placeholders
    const shapes = slide.getShapes();

    for (let shape of shapes) {
      try {
        const placeholder = shape.getPlaceholderType();

        if (placeholder === SlidesApp.PlaceholderType.CENTERED_TITLE && slideData.title) {
          shape.getText().setText(slideData.title);
          const titleStyle = shape.getText().getTextStyle();
          titleStyle.setFontSize(config['title_slide_font_size'] || 44);
          titleStyle.setFontFamily(config['header_font'] || 'Arial');
          titleStyle.setForegroundColor(config['text_color'] || '#000000');
          titleStyle.setBold(true);
        }

        if (placeholder === SlidesApp.PlaceholderType.SUBTITLE && slideData.subtitle) {
          shape.getText().setText(slideData.subtitle);
          const subtitleStyle = shape.getText().getTextStyle();
          subtitleStyle.setFontSize(config['subtitle_font_size'] || 28);
          subtitleStyle.setFontFamily(config['body_font'] || 'Arial');
          subtitleStyle.setForegroundColor(config['text_color'] || '#000000');
        }
      } catch (e) {
        // Shape might not be a placeholder, continue
        continue;
      }
    }
  } catch (e) {
    Logger.log('Error in createTitleSlide: ' + e.toString());
  }

  return slide;
}

function createSectionSlide(presentation, slideData, config) {
  // Use TITLE_AND_BODY layout instead of SECTION_HEADER for more reliable placeholder access
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);

  // Get theme (now always black on white)
  const theme = getThemeForSection(slideData.section_id, config);

  try {
    const shapes = slide.getShapes();
    let titleShape = null;
    let bodyShape = null;

    // Find title and body shapes
    for (let shape of shapes) {
      try {
        const placeholder = shape.getPlaceholderType();
        if (placeholder === SlidesApp.PlaceholderType.TITLE) {
          titleShape = shape;
        } else if (placeholder === SlidesApp.PlaceholderType.BODY) {
          bodyShape = shape;
        }
      } catch (e) {
        continue;
      }
    }

    // Set section title
    if (titleShape && slideData.title) {
      titleShape.getText().setText(slideData.title);
      const titleStyle = titleShape.getText().getTextStyle();
      titleStyle.setFontSize(config['section_slide_font_size'] || 40);
      titleStyle.setFontFamily(config['header_font'] || 'Arial');
      titleStyle.setForegroundColor(theme.textColor);
      titleStyle.setBold(true);

      // Center align the title
      titleShape.getText().getParagraphStyle()
        .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    }

    // Set section content in body
    if (bodyShape) {
      let bodyContent = '';

      // Add subtitle if present
      if (slideData.subtitle) {
        bodyContent = slideData.subtitle + '\n\n';
      }

      // Add bullets if present
      if (slideData.bullets) {
        bodyContent += formatBullets(slideData.bullets);
      }

      if (bodyContent.trim() !== '') {
        bodyShape.getText().setText(bodyContent);
        const bodyStyle = bodyShape.getText().getTextStyle();
        bodyStyle.setFontSize(config['subtitle_font_size'] || 24);
        bodyStyle.setFontFamily(config['body_font'] || 'Arial');
        bodyStyle.setForegroundColor(theme.textColor);

        // Center align the body content
        bodyShape.getText().getParagraphStyle()
          .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

        // Style subtitle differently if present
        if (slideData.subtitle) {
          try {
            const subtitleRange = bodyShape.getText().getRange(0, slideData.subtitle.length);
            subtitleRange.getTextStyle().setFontSize(config['subtitle_font_size'] || 24);
            subtitleRange.getTextStyle().setItalic(true);
          } catch (rangeError) {
            Logger.log('Could not style subtitle range: ' + rangeError.toString());
          }
        }
      }
    }

    Logger.log('Section slide created with title: ' + slideData.title);
  } catch (e) {
    Logger.log('Error in createSectionSlide: ' + e.toString());
  }

  return slide;
}

function createBasicSlide(presentation, slideData, config) {
  const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);

  // Set background for themed sections
  const theme = getThemeForSection(slideData.section_id, config);
  if (theme.backgroundColor && theme.backgroundColor !== '#FFFFFF') {
    try {
      slide.getBackground().getSolidFill().getColor().setRgbColor(theme.backgroundColor);
    } catch (bgError) {
      Logger.log('Could not set background: ' + bgError.toString());
    }
  }

  try {
    const shapes = slide.getShapes();
    let titleShape = null;
    let bodyShape = null;

    // Find title and body shapes
    for (let shape of shapes) {
      try {
        const placeholder = shape.getPlaceholderType();
        if (placeholder === SlidesApp.PlaceholderType.TITLE) {
          titleShape = shape;
        } else if (placeholder === SlidesApp.PlaceholderType.BODY) {
          bodyShape = shape;
        }
      } catch (e) {
        continue;
      }
    }

    // Set title
    if (titleShape && slideData.title) {
      titleShape.getText().setText(slideData.title);
      const titleStyle = titleShape.getText().getTextStyle();
      titleStyle.setFontSize(config['content_title_font_size'] || 36);
      titleStyle.setFontFamily(config['header_font'] || 'Arial');
      titleStyle.setForegroundColor(theme.textColor);
      titleStyle.setBold(true);
    }

    // Set body content
    if (bodyShape) {
      let bodyContent = '';

      // Add subtitle if present
      if (slideData.subtitle) {
        bodyContent = slideData.subtitle + '\n\n';
      }

      // Add bullets if present
      if (slideData.bullets) {
        bodyContent += formatBullets(slideData.bullets);
      }

      if (bodyContent.trim() !== '') {
        bodyShape.getText().setText(bodyContent);
        const bodyStyle = bodyShape.getText().getTextStyle();
        bodyStyle.setFontSize(config['bullet_font_size'] || 20);
        bodyStyle.setFontFamily(config['body_font'] || 'Arial');
        bodyStyle.setForegroundColor(theme.textColor);

        // Style subtitle differently if present
        if (slideData.subtitle) {
          try {
            const subtitleRange = bodyShape.getText().getRange(0, slideData.subtitle.length);
            subtitleRange.getTextStyle().setFontSize(config['subtitle_font_size'] || 24);
            subtitleRange.getTextStyle().setItalic(true);
          } catch (rangeError) {
            Logger.log('Could not style subtitle range: ' + rangeError.toString());
          }
        }
      }
    }

    // Add special elements - check for any chart_ref or media_ref regardless of layout
    if (slideData.chart_ref && slideData.chart_ref.trim() !== '') {
      addChartPlaceholder(slide, slideData.chart_ref, presentation);
    }

    if (slideData.media_ref && slideData.media_ref.trim() !== '') {
      addImagePlaceholder(slide, slideData.media_ref, presentation);
    }

  } catch (e) {
    Logger.log('Error in createBasicSlide: ' + e.toString());
  }

  return slide;
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

function formatBullets(bulletsText) {
  if (!bulletsText) return '';

  // Handle different bullet formats
  let bullets;
  if (bulletsText.includes(' • ')) {
    bullets = bulletsText.split(' • ');
  } else if (bulletsText.includes('|')) {
    bullets = bulletsText.split('|');
  } else {
    bullets = [bulletsText];
  }

  // Format as bullet list
  return bullets
    .filter(b => b.trim() !== '')
    .map(b => '• ' + b.trim())
    .join('\n');
}

function getThemeForSection(sectionId, config) {
  // Simplified to always use white background and black text
  return {
    backgroundColor: '#FFFFFF',
    textColor: '#000000'
  };
}

function addChartPlaceholder(slide, chartRef, presentation) {
  try {
    // Try to find and insert the actual chart image from Google Drive
    const chartName = chartRef.toString().trim();

    if (chartName) {
      try {
        // Search for the chart file in the images folder
        const folders = DriveApp.getFoldersByName('images');
        let chartFile = null;

        while (folders.hasNext() && !chartFile) {
          const folder = folders.next();
          const files = folder.getFilesByName(chartName);
          if (files.hasNext()) {
            chartFile = files.next();
            break;
          }
        }

        // If not found in images folder, search in root
        if (!chartFile) {
          const files = DriveApp.getFilesByName(chartName);
          if (files.hasNext()) {
            chartFile = files.next();
          }
        }

        if (chartFile) {
          // Insert the actual chart image
          const chart = slide.insertImage(chartFile.getBlob());

          // Get original dimensions to maintain aspect ratio
          const originalWidth = chart.getWidth();
          const originalHeight = chart.getHeight();
          const aspectRatio = originalWidth / originalHeight;

          // Calculate new size maintaining aspect ratio
          const maxWidth = presentation.getPageWidth() * 0.7;
          const maxHeight = presentation.getPageHeight() * 0.5;

          let newWidth, newHeight;
          if (maxWidth / maxHeight > aspectRatio) {
            // Height is limiting factor
            newHeight = maxHeight;
            newWidth = newHeight * aspectRatio;
          } else {
            // Width is limiting factor
            newWidth = maxWidth;
            newHeight = newWidth / aspectRatio;
          }

          // Center the chart
          const chartLeft = (presentation.getPageWidth() - newWidth) / 2;
          const chartTop = presentation.getPageHeight() * 0.3;

          chart.setWidth(newWidth);
          chart.setHeight(newHeight);
          chart.setLeft(chartLeft);
          chart.setTop(chartTop);

          Logger.log('Successfully inserted chart: ' + chartName);
          return;
        } else {
          Logger.log('Chart file not found in Drive: ' + chartName);
        }
      } catch (e) {
        Logger.log('Error loading chart from Drive: ' + e.toString());
      }
    }

    // Fall back to placeholder if chart not found
    const chartPlaceholder = slide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      presentation.getPageWidth() * 0.1,
      presentation.getPageHeight() * 0.4,
      presentation.getPageWidth() * 0.8,
      presentation.getPageHeight() * 0.4
    );

    chartPlaceholder.getText().setText('📊 CHART NOT FOUND: ' + chartRef);
    chartPlaceholder.getText().getTextStyle()
      .setFontSize(24)
      .setForegroundColor('#666666');
    chartPlaceholder.getText().getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    chartPlaceholder.getFill().setSolidFill('#F0F0F0');
    chartPlaceholder.getBorder().setWeight(2).getLineFill().setSolidFill('#CCCCCC');
  } catch (e) {
    Logger.log('Could not add chart placeholder: ' + e.toString());
  }
}

function addImagePlaceholder(slide, mediaRef, presentation) {
  try {
    // Try to find and insert the actual image from Google Drive
    const imageName = mediaRef.toString().trim();

    if (imageName) {
      try {
        // Search for the image file in the images folder
        const folders = DriveApp.getFoldersByName('images');
        let imageFile = null;

        while (folders.hasNext() && !imageFile) {
          const folder = folders.next();
          const files = folder.getFilesByName(imageName);
          if (files.hasNext()) {
            imageFile = files.next();
            break;
          }
        }

        // If not found in images folder, search in root
        if (!imageFile) {
          const files = DriveApp.getFilesByName(imageName);
          if (files.hasNext()) {
            imageFile = files.next();
          }
        }

        if (imageFile) {
          // Insert the actual image
          const image = slide.insertImage(imageFile.getBlob());

          // Get original dimensions to maintain aspect ratio
          const originalWidth = image.getWidth();
          const originalHeight = image.getHeight();
          const aspectRatio = originalWidth / originalHeight;

          // Calculate new size maintaining aspect ratio
          const maxWidth = presentation.getPageWidth() * 0.7;
          const maxHeight = presentation.getPageHeight() * 0.5;

          let newWidth, newHeight;
          if (maxWidth / maxHeight > aspectRatio) {
            // Height is limiting factor
            newHeight = maxHeight;
            newWidth = newHeight * aspectRatio;
          } else {
            // Width is limiting factor
            newWidth = maxWidth;
            newHeight = newWidth / aspectRatio;
          }

          // Center the image
          const imgLeft = (presentation.getPageWidth() - newWidth) / 2;
          const imgTop = presentation.getPageHeight() * 0.3;

          image.setWidth(newWidth);
          image.setHeight(newHeight);
          image.setLeft(imgLeft);
          image.setTop(imgTop);

          Logger.log('Successfully inserted image: ' + imageName);
          return;
        } else {
          Logger.log('Image file not found in Drive: ' + imageName);
        }
      } catch (e) {
        Logger.log('Error loading image from Drive: ' + e.toString());
      }
    }

    // Fall back to placeholder if image not found
    const imagePlaceholder = slide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      presentation.getPageWidth() * 0.1,
      presentation.getPageHeight() * 0.4,
      presentation.getPageWidth() * 0.8,
      presentation.getPageHeight() * 0.4
    );

    imagePlaceholder.getText().setText('🖼️ IMAGE NOT FOUND: ' + mediaRef);
    imagePlaceholder.getText().getTextStyle()
      .setFontSize(24)
      .setForegroundColor('#666666');
    imagePlaceholder.getText().getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    imagePlaceholder.getFill().setSolidFill('#E8E8E8');
    imagePlaceholder.getBorder().setWeight(2).getLineFill().setSolidFill('#CCCCCC');
  } catch (e) {
    Logger.log('Could not add image placeholder: ' + e.toString());
  }
}

function addFooterToAllSlides(presentation, config) {
  const footerText = config['footer_text'];
  if (!footerText) return;

  try {
    const slides = presentation.getSlides();
    const pageWidth = presentation.getPageWidth();
    const pageHeight = presentation.getPageHeight();

    slides.forEach((slide, index) => {
      // Skip title slide (first slide)
      if (index === 0) return;

      try {
        const footerBox = slide.insertTextBox(
          footerText,
          pageWidth * 0.05,
          pageHeight * 0.93,
          pageWidth * 0.9,
          pageHeight * 0.04
        );

        footerBox.getText().getTextStyle()
          .setFontSize(10)
          .setFontFamily(config['body_font'] || 'Arial')
          .setForegroundColor('#666666');
        footerBox.getText().getParagraphStyle()
          .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      } catch (footerError) {
        Logger.log(`Could not add footer to slide ${index + 1}: ${footerError.toString()}`);
      }
    });

    Logger.log('Added footer to slides');
  } catch (e) {
    Logger.log('Error adding footers: ' + e.toString());
  }
}

// ============================================================================
// DEBUG FUNCTIONS
// ============================================================================


function testSlideCreation() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = loadConfig(ss);
    const slides = loadSlides(ss);

    if (slides.length === 0) {
      SpreadsheetApp.getUi().alert('No slides found!');
      return;
    }

    // Create test presentation with timestamp in same folder as spreadsheet
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
    const testName = 'TEST_SingleSlide_' + timestamp;
    const presentation = SlidesApp.create(testName);

    // Move to same folder as spreadsheet
    const spreadsheetFile = DriveApp.getFileById(ss.getId());
    const folder = spreadsheetFile.getParents().next();
    const presentationFile = DriveApp.getFileById(presentation.getId());
    presentationFile.moveTo(folder);

    const defaultSlides = presentation.getSlides();
    if (defaultSlides.length > 0) {
      defaultSlides[0].remove();
    }

    // Create first slide
    const slideData = slides[0];
    Logger.log('Testing slide creation with data: ' + JSON.stringify(slideData));

    const slide = createSlide(presentation, slideData, config);

    if (slide) {
      const report =
        'Test slide created successfully!\n\n' +
        'File: ' + testName + '\n' +
        'Location: Same folder as this spreadsheet\n\n' +
        'Slide data:\n' +
        'Title: "' + slideData.title + '"\n' +
        'Layout: "' + slideData.layout + '"\n' +
        'Bullets: "' + (slideData.bullets ? slideData.bullets.substring(0, 100) + '...' : 'None') + '\n\n' +
        'Open presentation: ' + presentation.getUrl();

      SpreadsheetApp.getUi().alert('Test Results', report, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('Test failed: createSlide returned null');
    }

  } catch (error) {
    Logger.log('Test error: ' + error.toString());
    SpreadsheetApp.getUi().alert('Test failed: ' + error.toString());
  }
}


// ============================================================================
// SETUP FUNCTIONS
// ============================================================================


function testConfig() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = loadConfig(ss);

    const summary =
      'Configuration loaded successfully!\n\n' +
      'Title: ' + (config['deck_title'] || 'Not set') + '\n' +
      'Presenter: ' + (config['presenter_name'] || 'Not set') + '\n' +
      'Font sizes: Title=' + (config['title_slide_font_size'] || 'default') +
      ', Section=' + (config['section_slide_font_size'] || 'default') +
      ', Content=' + (config['content_title_font_size'] || 'default') + '\n' +
      'Total settings: ' + Object.keys(config).length;

    SpreadsheetApp.getUi().alert('Configuration Test', summary, SpreadsheetApp.getUi().ButtonSet.OK);

    Logger.log('Full config: ' + JSON.stringify(config, null, 2));

  } catch (error) {
    SpreadsheetApp.getUi().alert('Config test failed: ' + error.toString());
  }
}

function validateSlideData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const slides = loadSlides(ss);
    const issues = [];

    slides.forEach((slide, index) => {
      if (!slide.title) {
        issues.push(`Slide ${index + 1}: Missing title`);
      }
      if (slide.layout === 'Chart' && !slide.chart_ref) {
        issues.push(`Slide ${index + 1}: Chart slide missing chart reference`);
      }
      if (slide.layout === 'Image' && !slide.media_ref) {
        issues.push(`Slide ${index + 1}: Image slide missing media reference`);
      }
    });

    if (issues.length > 0) {
      SpreadsheetApp.getUi().alert('Validation Results',
        'Found ' + issues.length + ' issues:\n\n' + issues.slice(0, 8).join('\n') +
        (issues.length > 8 ? '\n\n...and ' + (issues.length - 8) + ' more' : ''),
        SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('Validation Passed',
        'All ' + slides.length + ' slides have required data!',
        SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Validation error: ' + error.toString());
  }
}

// ============================================================================
// AUTO-UPDATE FUNCTIONS (NON-DESTRUCTIVE WITH VERSIONING)
// ============================================================================

// Version management configuration
const GITHUB_CONFIG = {
  username: 'willscott-v2',
  repo: 'smx-oct-slides',
  branch: 'main',  // or 'v1.0', 'stable', etc.
  currentVersion: '1.2.1'  // Stored locally, compared with GitHub
};

// Helper function for robust fetching with retry logic
function fetchWithRetry(url, maxRetries = 3) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      console.log(`Fetching ${url} (attempt ${attempt}/${maxRetries})`);

      const response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: {
          'User-Agent': 'SMX-GEO-Generator/1.0.0',
          'Accept': 'text/plain,application/json,*/*'
        }
      });

      const responseCode = response.getResponseCode();
      console.log(`Response code: ${responseCode}`);

      if (responseCode === 200) {
        return response.getContentText();
      } else {
        throw new Error(`HTTP ${responseCode}: ${response.getContentText()}`);
      }
    } catch (error) {
      console.error(`Attempt ${attempt} failed:`, error.toString());
      if (attempt === maxRetries) {
        throw new Error(`Failed after ${maxRetries} attempts: ${error.toString()}`);
      }
      // Wait before retry (exponential backoff)
      Utilities.sleep(1000 * attempt);
    }
  }
}

function previewUpdatesFromGitHub() {
  try {
    // GitHub URLs with version support
    const baseUrl = `https://raw.githubusercontent.com/${GITHUB_CONFIG.username}/${GITHUB_CONFIG.repo}/${GITHUB_CONFIG.branch}`;
    const VERSION_URL = `${baseUrl}/version.json`;
    const CONFIG_URL = `${baseUrl}/config.csv`;
    const SLIDES_URL = `${baseUrl}/slides.csv`;

    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Show preview dialog
    const response = ui.alert('Preview Updates',
      'This will safely compare your current data with the latest from GitHub.\n\n' +
      'No changes will be made to your sheets.\n\n' +
      'Continue with preview?',
      ui.ButtonSet.YES_NO);

    if (response !== ui.Button.YES) return;

    let report = 'VERSION & UPDATE PREVIEW\n' + '='.repeat(50) + '\n\n';

    // Check version first
    let githubVersion, hasUpdates = false;
    try {
      // Fetch real version from GitHub with retry logic
      const versionContent = fetchWithRetry(VERSION_URL);
      githubVersion = JSON.parse(versionContent);

      report += `📦 VERSION CHECK:\n`;
      report += `   Current:  v${GITHUB_CONFIG.currentVersion}\n`;
      report += `   GitHub:   v${githubVersion.version}\n`;

      if (compareVersions(GITHUB_CONFIG.currentVersion, githubVersion.version) < 0) {
        hasUpdates = true;
        report += `   Status:   🆕 Update available!\n`;
        report += `   Released: ${githubVersion.releaseDate}\n\n`;

        report += `📝 WHAT'S NEW IN v${githubVersion.version}:\n`;
        githubVersion.changes.forEach(change => {
          report += `   • ${change}\n`;
        });
        report += '\n';
      } else {
        report += `   Status:   ✅ You have the latest version\n\n`;
      }

    } catch (versionError) {
      report += `📦 VERSION CHECK: ⚠️ Could not check version (${versionError.toString()})\n\n`;
      // Fall back to checking files directly
      githubVersion = { version: 'unknown', changes: ['Could not fetch version info'] };
    }

    if (hasUpdates) {
      // Real data comparison for updates
      try {
        const currentConfig = loadConfig(ss);
        const currentSlides = loadSlides(ss);

        // Fetch actual data from GitHub for comparison with retry logic
        const configContent = fetchWithRetry(CONFIG_URL);
        const slidesContent = fetchWithRetry(SLIDES_URL);

        const githubConfigData = Utilities.parseCsv(configContent);
        const githubSlidesData = Utilities.parseCsv(slidesContent, '|');

        report += `📊 DATA COMPARISON:\n`;
        report += `   Config:  ${Object.keys(currentConfig).length} settings (local) vs ${githubConfigData.length - 1} (GitHub)\n`;
        report += `   Slides:  ${currentSlides.length} slides (local) vs ${githubSlidesData.length - 1} (GitHub)\n`;

        if (githubConfigData.length - 1 !== Object.keys(currentConfig).length ||
            githubSlidesData.length - 1 !== currentSlides.length) {
          report += `   Status:  🔄 Data differences detected\n\n`;
        } else {
          report += `   Status:  ✅ Data appears current\n\n`;
        }

        report += `🎯 RECOMMENDED ACTION:\n`;
        report += `   1. Use "Apply Updates" to get v${githubVersion.version}\n`;
        report += `   2. Your data will be backed up automatically\n`;
        report += `   3. New features will be available immediately\n\n`;
      } catch (dataError) {
        report += `📊 DATA COMPARISON: ⚠️ Could not compare data\n\n`;
      }
    } else {
      report += `✅ NO UPDATES NEEDED\n`;
      report += `   You're running the latest version!\n\n`;
    }

    report += `🟢 LIVE MODE ACTIVE\n`;
    report += `   Connected to: ${baseUrl}\n`;
    report += `   Ready for real updates!\n`;

    // Show the preview report
    ui.alert('Preview Results', report, ui.ButtonSet.OK);
    Logger.log('Preview Report:\n' + report);

  } catch (error) {
    Logger.log('Preview error: ' + error.toString());
    SpreadsheetApp.getUi().alert('Preview failed: ' + error.toString());
  }
}

function applyUpdatesFromGitHub() {
  try {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Confirmation dialog
    const response = ui.alert('Apply Updates',
      'This will update your sheets with data from GitHub.\n\n' +
      '⚠️  Your current data will be backed up first.\n\n' +
      'Are you sure you want to continue?',
      ui.ButtonSet.YES_NO);

    if (response !== ui.Button.YES) return;

    // Create backup sheets first
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
    createBackupSheets(ss, timestamp);

    // Apply real updates from GitHub
    try {
      const baseUrl = `https://raw.githubusercontent.com/${GITHUB_CONFIG.username}/${GITHUB_CONFIG.repo}/${GITHUB_CONFIG.branch}`;
      const CONFIG_URL = `${baseUrl}/config.csv`;
      const SLIDES_URL = `${baseUrl}/slides.csv`;
      const VERSION_URL = `${baseUrl}/version.json`;

      // Fetch data from GitHub with retry logic
      const configContent = fetchWithRetry(CONFIG_URL);
      const slidesContent = fetchWithRetry(SLIDES_URL);
      const versionContent = fetchWithRetry(VERSION_URL);

      const configData = Utilities.parseCsv(configContent);
      const slidesData = Utilities.parseCsv(slidesContent, '|');
      const versionData = JSON.parse(versionContent);

      // Update Config sheet
      const configSheet = ss.getSheetByName('Config');
      if (configSheet) {
        configSheet.clear();
        configSheet.getRange(1, 1, configData.length, configData[0].length).setValues(configData);
      }

      // Update Slides sheet
      const slidesSheet = ss.getSheetByName('Slides');
      if (slidesSheet) {
        slidesSheet.clear();
        slidesSheet.getRange(1, 1, slidesData.length, slidesData[0].length).setValues(slidesData);
      }

      // Update local version
      setCurrentVersion(versionData.version);

      ui.alert('✅ UPDATE SUCCESSFUL!',
        `Updated to v${versionData.version}\n\n` +
        `📋 Changes applied:\n` +
        `   • Config: ${configData.length - 1} settings\n` +
        `   • Slides: ${slidesData.length - 1} slides\n\n` +
        `💾 Backup created: ${timestamp}\n\n` +
        `🎯 What's new:\n${versionData.changes.slice(0, 3).map(c => `   • ${c}`).join('\n')}`,
        ui.ButtonSet.OK);

      Logger.log(`Update applied successfully to v${versionData.version}`);

    } catch (updateError) {
      ui.alert('❌ UPDATE FAILED',
        `Could not fetch updates from GitHub:\n\n${updateError.toString()}\n\n` +
        `✅ Your data is safe - backups created: ${timestamp}\n\n` +
        `Check your internet connection and GitHub repository.`,
        ui.ButtonSet.OK);
      Logger.log('Update failed: ' + updateError.toString());
    }

  } catch (error) {
    Logger.log('Apply error: ' + error.toString());
    SpreadsheetApp.getUi().alert('Apply failed: ' + error.toString());
  }
}

function createBackupSheets(ss, timestamp) {
  try {
    // Backup Config sheet
    const configSheet = ss.getSheetByName('Config');
    if (configSheet) {
      const backupConfig = configSheet.copyTo(ss);
      backupConfig.setName(`Config_Backup_${timestamp}`);
    }

    // Backup Slides sheet
    const slidesSheet = ss.getSheetByName('Slides');
    if (slidesSheet) {
      const backupSlides = slidesSheet.copyTo(ss);
      backupSlides.setName(`Slides_Backup_${timestamp}`);
    }

    Logger.log(`Backup sheets created with timestamp: ${timestamp}`);

  } catch (error) {
    Logger.log('Backup error: ' + error.toString());
    throw error;
  }
}

function compareSheetData(currentData, newData, dataType) {
  const changes = [];

  if (currentData.length !== newData.length) {
    changes.push(`${dataType} count changed: ${currentData.length} → ${newData.length}`);
  }

  // Add more comparison logic here
  // For now, just basic length comparison

  return changes;
}

function compareVersions(version1, version2) {
  // Semantic version comparison (e.g., "1.0.0" vs "1.1.0")
  const v1parts = version1.split('.').map(Number);
  const v2parts = version2.split('.').map(Number);

  for (let i = 0; i < Math.max(v1parts.length, v2parts.length); i++) {
    const v1part = v1parts[i] || 0;
    const v2part = v2parts[i] || 0;

    if (v1part < v2part) return -1;
    if (v1part > v2part) return 1;
  }

  return 0;
}

function getCurrentVersion() {
  // Get current version from PropertiesService (persistent storage)
  const userProperties = PropertiesService.getUserProperties();
  return userProperties.getProperty('GEO_GENERATOR_VERSION') || GITHUB_CONFIG.currentVersion;
}

function setCurrentVersion(version) {
  // Store current version in PropertiesService
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('GEO_GENERATOR_VERSION', version);
}