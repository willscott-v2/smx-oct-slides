/**
 * Complete GEO Master Class Presentation Generator
 * Version: 4.0 - Simplified and debugged
 */

// ============================================================================
// MENU & INITIALIZATION
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GEO Presentation')
    .addItem('Generate Slides', 'generatePresentation')
    .addItem('Test Configuration', 'testConfig')
    .addItem('Validate Data', 'validateSlideData')
    .addItem('Setup Config Tab', 'setupConfigTab')
    .addSeparator()
    .addItem('🔍 Debug Slide Data', 'debugSlideData')
    .addItem('🔍 Test One Slide', 'testSlideCreation')
    .addItem('🔍 Check Headers', 'checkSlideHeaders')
    .addSeparator()
    .addItem('🔄 Preview Updates (Safe)', 'previewUpdatesFromGitHub')
    .addItem('📥 Apply Updates', 'applyUpdatesFromGitHub')
    .addToUi();
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

    // Add special elements for specific layouts
    if (slideData.layout === 'Chart' && slideData.chart_ref) {
      addChartPlaceholder(slide, slideData.chart_ref, presentation);
    }

    if (slideData.layout === 'Image' && slideData.media_ref) {
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
    const chartPlaceholder = slide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      presentation.getPageWidth() * 0.1,
      presentation.getPageHeight() * 0.4,
      presentation.getPageWidth() * 0.8,
      presentation.getPageHeight() * 0.4
    );

    chartPlaceholder.getText().setText('📊 CHART: ' + chartRef);
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
    const imagePlaceholder = slide.insertShape(
      SlidesApp.ShapeType.RECTANGLE,
      presentation.getPageWidth() * 0.1,
      presentation.getPageHeight() * 0.4,
      presentation.getPageWidth() * 0.8,
      presentation.getPageHeight() * 0.4
    );

    imagePlaceholder.getText().setText('🖼️ IMAGE: ' + mediaRef);
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

function debugSlideData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const slides = loadSlides(ss);

    Logger.log('=== SLIDE DATA DEBUG ===');
    Logger.log('Total slides loaded: ' + slides.length);

    // Show first 3 slides with all their data
    for (let i = 0; i < Math.min(3, slides.length); i++) {
      Logger.log('--- Slide ' + (i + 1) + ' ---');
      Logger.log('Order: ' + slides[i].order);
      Logger.log('Layout: "' + slides[i].layout + '"');
      Logger.log('Title: "' + slides[i].title + '"');
      Logger.log('Subtitle: "' + slides[i].subtitle + '"');
      Logger.log('Bullets: "' + slides[i].bullets.substring(0, 100) + '..."');
      Logger.log('Speaker Notes: "' + slides[i].speaker_notes.substring(0, 100) + '..."');
      Logger.log('Section ID: "' + slides[i].section_id + '"');
    }

    // Check for common issues
    const issues = [];
    slides.forEach((slide, index) => {
      if (!slide.title || slide.title === '') {
        issues.push(`Slide ${index + 1}: Empty title`);
      }
      if (!slide.bullets || slide.bullets === '') {
        issues.push(`Slide ${index + 1}: Empty bullets`);
      }
    });

    Logger.log('=== ISSUES FOUND ===');
    Logger.log('Issues: ' + issues.length);
    issues.forEach(issue => Logger.log('- ' + issue));

    const summary =
      'DEBUG RESULTS:\n\n' +
      'Slides loaded: ' + slides.length + '\n' +
      'Issues found: ' + issues.length + '\n\n' +
      'Check Apps Script logs for detailed data.';

    SpreadsheetApp.getUi().alert('Debug Results', summary, SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    Logger.log('DEBUG ERROR: ' + error.toString());
    SpreadsheetApp.getUi().alert('Debug failed: ' + error.toString());
  }
}

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

function checkSlideHeaders() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Slides');

    if (!sheet) {
      SpreadsheetApp.getUi().alert('No Slides sheet found!');
      return;
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    Logger.log('=== HEADERS DEBUG ===');
    Logger.log('Found headers: ' + headers.length);
    headers.forEach((header, index) => {
      Logger.log('Column ' + (index + 1) + ': "' + header + '"');
    });

    // Check for required headers
    const required = ['order', 'title'];
    const missing = [];

    required.forEach(req => {
      const found = headers.some(h => h.toString().toLowerCase().replace(/\s+/g, '_') === req);
      if (!found) {
        missing.push(req);
      }
    });

    let message = 'Headers found: ' + headers.length + '\n\n';
    message += 'Headers: ' + headers.join(', ') + '\n\n';
    if (missing.length > 0) {
      message += 'MISSING REQUIRED: ' + missing.join(', ');
    } else {
      message += 'All required headers present!';
    }

    SpreadsheetApp.getUi().alert('Header Check', message, SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    SpreadsheetApp.getUi().alert('Header check failed: ' + error.toString());
  }
}

// ============================================================================
// SETUP FUNCTIONS
// ============================================================================

function setupConfigTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('Config');

  if (configSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('Config tab exists',
      'Do you want to replace it with default settings?',
      ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) {
      return;
    }
    ss.deleteSheet(configSheet);
  }

  // Create new config sheet
  configSheet = ss.insertSheet('Config');

  // Enhanced configuration with all needed settings
  const configData = [
    ['Setting', 'Value'],
    ['Deck Title', 'Generative Engine Optimization Master Class'],
    ['Deck Subtitle', 'Online October 7, 2025'],
    ['Presenter Name', 'Will Scott'],
    ['Company Name', 'Search Influence'],
    ['Footer Text', 'willscott.me/links  Search Influence © 2025'],
    ['Header Font', 'Arial'],
    ['Body Font', 'Arial'],
    ['Title Slide Font Size', 44],
    ['Section Slide Font Size', 40],
    ['Content Title Font Size', 36],
    ['Subtitle Font Size', 24],
    ['Bullet Font Size', 20],
    ['Speaker Notes Font Size', 14],
    ['Text Color', '#000000'],
    ['Highlight Color', '#E8E8E8'],
    ['Session 1 Background', '#667eea'],
    ['Session 1 Text Color', '#FFFFFF'],
    ['Session 2 Background', '#84fab0'],
    ['Session 2 Text Color', '#000000'],
    ['Session 3 Background', '#a8edea'],
    ['Session 3 Text Color', '#000000'],
    ['Session 4 Background', '#ffecd2'],
    ['Session 4 Text Color', '#000000'],
    ['Recap Background', '#F5F5F5'],
    ['Recap Text Color', '#333333'],
    ['QR Code Size', 150]
  ];

  // Write configuration data
  const range = configSheet.getRange(1, 1, configData.length, 2);
  range.setValues(configData);

  // Format header row
  configSheet.getRange(1, 1, 1, 2)
    .setFontWeight('bold')
    .setBackground('#E0E0E0');

  // Auto-resize columns
  configSheet.autoResizeColumns(1, 2);

  SpreadsheetApp.getUi().alert('Config tab created successfully!');
}

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
  currentVersion: '1.0.0'  // Stored locally, compared with GitHub
};

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
      // In simulation mode, create a mock version response
      // In production: const versionResponse = UrlFetchApp.fetch(VERSION_URL);
      githubVersion = {
        version: '1.1.0',
        releaseDate: '2024-10-03',
        changes: [
          'Added automated update functionality',
          'Improved slide formatting consistency',
          'Enhanced error handling'
        ]
      };

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
      report += `📦 VERSION CHECK: ⚠️ Could not check version\n\n`;
    }

    if (hasUpdates) {
      // Simulate data comparison for updates
      const currentConfig = loadConfig(ss);
      const currentSlides = loadSlides(ss);

      report += `📊 DATA COMPARISON:\n`;
      report += `   Config:  ${Object.keys(currentConfig).length} settings\n`;
      report += `   Slides:  ${currentSlides.length} slides\n`;
      report += `   Status:  🔄 Updates detected\n\n`;

      report += `🎯 RECOMMENDED ACTION:\n`;
      report += `   1. Use "Apply Updates" to get v${githubVersion.version}\n`;
      report += `   2. Your data will be backed up automatically\n`;
      report += `   3. New features will be available immediately\n\n`;
    } else {
      report += `✅ NO UPDATES NEEDED\n`;
      report += `   You're running the latest version!\n\n`;
    }

    report += `🚧 SIMULATION MODE ACTIVE\n`;
    report += `   Update GitHub URLs in code to enable live updates\n`;
    report += `   Test mode prevents accidental changes\n`;

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

    // For testing, show what would happen
    ui.alert('Backup Created',
      `✅ Backup sheets created with timestamp: ${timestamp}\n\n` +
      '🚧 SIMULATION MODE:\n' +
      'In production, this would:\n' +
      '1. Fetch latest data from GitHub\n' +
      '2. Update your Config and Slides sheets\n' +
      '3. Show a summary of changes\n\n' +
      'To enable, update the GitHub URLs in the code.',
      ui.ButtonSet.OK);

    Logger.log('Update applied (simulation mode)');

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