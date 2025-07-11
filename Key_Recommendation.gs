function updateLocationsWithCounts() {
  const presentationId = '1oN9JtyxZfUl4wKrhPCifIRi6lw2Hked1DSwYHsjxwKE';
  const presentation = SlidesApp.openById(presentationId);
  const slides = presentation.getSlides();

  // Extract district name from the first slide
  const district = extractDistrictFromFirstSlide(slides[0]);
  Logger.log(`Extracted District: ${district}`);

  // Spreadsheet URLs
  const criticalCorridorSheetUrl = 'https://docs.google.com/spreadsheets/d/1KAhWeR-xi7T19T8nFUO1NTTjlIpbqtLkGC1RqSk8T3g/edit';
  const criticalLocationsSheetUrl = 'https://docs.google.com/spreadsheets/d/1eJrjgJD4kSWo_3rtoyNkSE0ZivE_CQGl9zDnYbolX_4/edit';

  const criticalCorridorSheet = SpreadsheetApp.openByUrl(criticalCorridorSheetUrl).getSheets()[0];
  const criticalLocationsSheet = SpreadsheetApp.openByUrl(criticalLocationsSheetUrl).getSheets()[0];

  // Google Drive folder ID
  const driveFolderId = 'DRIVE_LINK_WHERE_FILES_ARE_TO_BE_SAVED';
  const driveFolder = DriveApp.getFolderById(driveFolderId);

  // Log sheet names to verify
  Logger.log(`Critical Corridor Sheet Name: ${criticalCorridorSheet.getName()}`);
  Logger.log(`Critical Locations Sheet Name: ${criticalLocationsSheet.getName()}`);

  const targetTitles = [
    "Key Recommendations: NHAI",
    "Key Recommendations: PWD",
    "Key Recommendations: Urban/Local Body"
  ];

  slides.forEach(slide => {
    const pageElements = slide.getPageElements();
    let agency = '';

    for (let element of pageElements) {
      if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
        const text = element.asShape().getText().asString();

        for (let title of targetTitles) {
          if (text.includes(title)) {
            agency = title.split(':')[1].trim();
            Logger.log(`=== Processing Slide for Agency: ${agency} ===`);

            const tables = slide.getTables();
            for (const table of tables) {
              const numRows = table.getNumRows();

              for (let row = 0; row < numRows; row++) {
                try {
                  const cell = table.getCell(row, 1);
                  const cellText = cell.getText().asString().trim();

                  // Critical Corridor Check
                  if (cellText.toLowerCase().includes("critical corridor")) {
                    Logger.log(`Found Critical Corridor Row: "${cellText}"`);
                    const { count, filteredRows } = getFilteredCount(criticalCorridorSheet, district, agency, 'critical corridor');
                    Logger.log(`Filtered Critical Corridors Count: ${count}`);

                    // Create a new spreadsheet for Critical Corridors
                    const headers = criticalCorridorSheet.getRange(1, 1, 1, criticalCorridorSheet.getLastColumn()).getValues()[0];
                    const spreadsheetName = `${agency}_${district}_Critical_Corridors`;
                    const spreadsheetUrl = createSpreadsheetInFolder(driveFolder, spreadsheetName, headers, filteredRows);

                    updateCellWithCount(table, row, count, spreadsheetUrl, slide, agency, 'Critical Corridor');
                  }

                  // Crash Type Checks
                  else {
                    const crashTypes = [
                      "Head-on",
                      "Rear End",
                      "Side-Impact",
                      "Night time",
                      "Pedestrian"
                    ];

                    for (let crashType of crashTypes) {
                      if (cellText.toLowerCase().includes(crashType.toLowerCase())) {
                        Logger.log(`Found Crash Type Row: "${cellText}" (Type: ${crashType})`);
                        const { count, filteredRows } = getFilteredCount(criticalLocationsSheet, district, agency, crashType);
                        Logger.log(`Filtered ${crashType} Count: ${count}`);

                        // Create a new spreadsheet for Crash Types
                        const headers = criticalLocationsSheet.getRange(1, 1, 1, criticalLocationsSheet.getLastColumn()).getValues()[0];
                        const spreadsheetName = `${agency}_${district}_${crashType}`;
                        const spreadsheetUrl = createSpreadsheetInFolder(driveFolder, spreadsheetName, headers, filteredRows);

                        updateCellWithCount(table, row, count, spreadsheetUrl, slide, agency, crashType);
                      }
                    }
                  }
                } catch (e) {
                  Logger.log(`Error reading cell at row ${row}: ${e.message}`);
                  continue;
                }
              }
            }
          }
        }
      }
    }
  });

  Logger.log("=== Update Completed ===");
}

// Extract district name from the first slide
function extractDistrictFromFirstSlide(slide) {
  const pageElements = slide.getPageElements();
  let district = null;

  for (let element of pageElements) {
    if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      const text = element.asShape().getText().asString().trim();
      Logger.log(`First Slide Shape Text: "${text}"`);

      // Look for common district name patterns
      // 1. "District: Mandya" or "District: Mandya, Karnataka"
      if (text.toLowerCase().includes("district:")) {
        const match = text.match(/District:\s*([A-Za-z]+)(?:,.*)?/i);
        if (match) {
          district = match[1].trim();
          break;
        }
      }
      // 2. "Mandya District"
      else if (text.toLowerCase().includes("district")) {
        const match = text.match(/([A-Za-z]+)\s+District/i);
        if (match) {
          district = match[1].trim();
          break;
        }
      }
      // 3. "Mandya-Karnataka" (new pattern)
      else if (text.includes("-Karnataka")) {
        const match = text.match(/\b([A-Za-z]+)-Karnataka\b/i);
        if (match) {
          district = match[1].trim();
          break;
        }
      }
      // 4. "Mandya, Karnataka" or standalone "Mandya"
      else {
        const match = text.match(/\b([A-Za-z]+)(?:,\s*[A-Za-z]+)?\b/);
        if (match) {
          // Avoid matching generic words; check if the word is a potential district name
          const potentialDistrict = match[1].trim();
          // List of known Karnataka districts for validation (as of 2025)
          const karnatakaDistricts = [
            "Bagalkot", "Ballari", "Belagavi", "Bengaluru Urban", "Bengaluru Rural", "Bidar",
            "Chamarajanagar", "Chikkaballapur", "Chikkamagaluru", "Chitradurga", "Dakshina Kannada",
            "Davanagere", "Dharwad", "Gadag", "Hassan", "Haveri", "Kalaburagi", "Kodagu",
            "Kolar", "Koppal", "Mandya", "Mysuru", "Raichur", "Ramanagara", "Shivamogga",
            "Tumakuru", "Udupi", "Uttara Kannada", "Vijayapura", "Yadgir"
          ];
          if (karnatakaDistricts.includes(potentialDistrict)) {
            district = potentialDistrict;
            break;
          }
        }
      }
    }
  }

  if (!district) {
    Logger.log("Error: Could not extract district name from the first slide.");
    throw new Error("District name not found on the first slide. Please ensure the slide contains the district name.");
  }

  return district;
}

// Count matching rows and return both the count and the filtered rows
function getFilteredCount(sheet, district, agency, type) {
  const data = sheet.getDataRange().getValues();
  Logger.log(`Filtering for: District="${district}", Agency="${agency}", Type="${type}"`);

  const isCriticalCorridorSheet = sheet.getName().toLowerCase().includes("corridor");
  let filtered;

  if (isCriticalCorridorSheet) {
    // Critical Corridor Spreadsheet: District (A=0), Agency (E=4)
    filtered = data.filter((row, index) => {
      if (index === 0) return false; // Skip header
      const rowDistrict = row[0] ? row[0].toString().trim() : ''; // Column A
      const rowAgency = row[4] ? row[4].toString().trim() : '';   // Column E

      // Extract district name before the comma (e.g., "Mandya, Karnataka" -> "Mandya")
      const districtName = rowDistrict.split(',')[0].trim();

      // Normalize agency by removing "State" and trimming (e.g., "State PWD" -> "PWD")
      const normalizedRowAgency = rowAgency.toLowerCase().replace('state', '').trim();
      const normalizedAgency = agency.toLowerCase().trim();

      const match = districtName.toLowerCase() === district.toLowerCase() &&
                    normalizedRowAgency === normalizedAgency;
      return match;
    });
  } else {
    // Critical Locations Spreadsheet: District (B=1), Agency (F=5), Type of Crashes (C=2)
    // Map slide crash types to spreadsheet crash types
    const typeMapping = {
      "Head-on": "Head-on Collision",
      "Rear End": "Rear-end",
      "Side-Impact": "Side-Impact",
      "Night time": "Night Time",
      "Pedestrian": "Pedestrain" // Account for spreadsheet misspelling
    };
    const spreadsheetType = typeMapping[type] || type;

    filtered = data.filter((row, index) => {
      if (index === 0) return false; // Skip header
      const rowDistrict = row[1] ? row[1].toString().trim() : ''; // Column B
      const rowAgency = row[5] ? row[5].toString().trim() : '';   // Column F
      const rowType = row[2] ? row[2].toString().trim() : '';     // Column C

      // Normalize agency name by removing spaces and slashes
      const normalizedRowAgency = rowAgency.toLowerCase().replace(/[\s\/]/g, '');
      const normalizedAgency = agency.toLowerCase().replace(/[\s\/]/g, '');

      const match = rowDistrict.toLowerCase() === district.toLowerCase() &&
                    normalizedRowAgency === normalizedAgency &&
                    rowType && rowType.toLowerCase() === spreadsheetType.toLowerCase();
      return match;
    });
  }

  Logger.log(`Matched Rows: ${filtered.length}`);
  return { count: filtered.length, filteredRows: filtered };
}

// Create a new spreadsheet in the specified Drive folder with the filtered data and return its URL
function createSpreadsheetInFolder(folder, spreadsheetName, headers, filteredRows) {
  try {
    // Check if a spreadsheet with the same name already exists in the folder
    const files = folder.getFilesByName(spreadsheetName);
    let spreadsheet;

    if (files.hasNext()) {
      // If it exists, open it and clear the existing data
      spreadsheet = SpreadsheetApp.openById(files.next().getId());
      const sheet = spreadsheet.getSheets()[0];
      sheet.clear(); // Clear existing data
    } else {
      // If it doesn't exist, create a new spreadsheet
      spreadsheet = SpreadsheetApp.create(spreadsheetName);
      // Move the spreadsheet to the specified folder
      const file = DriveApp.getFileById(spreadsheet.getId());
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file); // Remove from root
    }

    // Get the first sheet of the spreadsheet
    const sheet = spreadsheet.getSheets()[0];

    // Write the headers
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Write the filtered data (if any)
    if (filteredRows.length > 0) {
      sheet.getRange(2, 1, filteredRows.length, headers.length).setValues(filteredRows);
    }

    Logger.log(`Created/Updated spreadsheet: ${spreadsheetName} with ${filteredRows.length} rows`);
    return spreadsheet.getUrl(); // Return the URL of the spreadsheet
  } catch (e) {
    Logger.log(`Error creating spreadsheet ${spreadsheetName}: ${e.message}`);
    return null; // Return null if there's an error
  }
}

// Helper function to append URL to slide's speaker notes
function appendUrlToSlideNotes(slide, agency, type, spreadsheetUrl) {
  if (!spreadsheetUrl) {
    Logger.log(`No URL provided for ${type} for ${agency}`);
    return;
  }

  try {
    const notesPage = slide.getNotesPage();
    const speakerNotesShape = notesPage.getSpeakerNotesShape();
    const notesText = speakerNotesShape.getText();
    const noteText = `${type} for ${agency}: ${spreadsheetUrl}\n`;
    notesText.appendText(noteText);
    Logger.log(`Appended to slide notes: "${noteText.trim()}"`);
  } catch (e) {
    Logger.log(`Error appending to speaker notes for ${type} for ${agency}: ${e.message}`);
  }
}

// Replace the number in a cell with the updated count and add the URL to slide notes
function updateCellWithCount(table, rowIndex, count, spreadsheetUrl, slide, agency, type) {
  try {
    const cell = table.getCell(rowIndex, 1);
    const textRange = cell.getText();
    const currentText = textRange.asString();

    // Find the position of the first number in the text
    const match = currentText.match(/\d+/);
    if (match) {
      // Replace the number with the new count
      const newText = currentText.replace(/\d+/, count);
      textRange.setText(newText);

      // Add the spreadsheet URL to the slide's speaker notes
      appendUrlToSlideNotes(slide, agency, type, spreadsheetUrl);
    } else {
      Logger.log(`No number found in cell text: "${currentText}"`);
    }

    Logger.log(`Updated Text: "${currentText}" => "${textRange.asString()}"`);
  } catch (e) {
    Logger.log(`Error updating cell at row ${rowIndex}: ${e.message}`);
  }
}