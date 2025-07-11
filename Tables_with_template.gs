function insertCriticalLocationTables() {
  const presentation = SlidesApp.getActivePresentation(); //access the currently active/open Google Slides presentation in your browser

  //you can change this piece of code with the one that accesses the presentation with the URL

  const slides = presentation.getSlides();

  // Step 1: Extract district name
  let districtName = null;
  const firstSlide = slides[0];
  for (const shape of firstSlide.getShapes()) {
    if (shape.getText) {
      const text = shape.getText().asString().trim();
      const lines = text.split(/\r?\n/).map(line => line.trim()).filter(line => line.length > 0);
      for (const line of lines) {
        if (line.toLowerCase().includes("road safety")) continue;
        if (line.includes("-")) {
          districtName = line.split("-")[0].trim();
          break;
        }
        if (/^[A-Z][a-z]+(?:\s[A-Z][a-z]+)?$/.test(line)) {
          districtName = line;
          break;
        }
      }
      if (districtName) break;
    }
  }

  // Step 2: Find the template slide
  let templateSlide = null;
  for (const slide of slides) {
    for (const shape of slide.getShapes()) {
      if (shape.getText && shape.getText().asString().includes("Critical Locations: Head-on Collisions")) {
        templateSlide = slide;
        break;
      }
    }
    if (templateSlide) break;
  }

  if (!templateSlide) {
    Logger.log("Template slide not found.");
    return;
  }

  const sheet = SpreadsheetApp.openById("SPREADSHEET_ID_FROM_URL")
    .getSheetByName("Combined Data for Presentation");

  if (!sheet) {
    Logger.log("Sheet 'Combined Data for Presentation' not found.");
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const districtCol = headers.indexOf("District");
  const crashTypeCol = headers.indexOf("Type of Crashes");
  const fatalitiesCol = headers.indexOf("Fatalities");

  if (districtCol === -1 || crashTypeCol === -1 || fatalitiesCol === -1) {
    Logger.log("Required columns not found.");
    return;
  }

  const selectedColumns = [
    "Rank", "Road Name", "Agency Name",
    "Co-ordinates \n(Lat, Long)", "Location Name", "Fatal Crashes", "Fatalities"
  ];
  const columnIndexes = selectedColumns.map(col => headers.indexOf(col));
  const rankColIndex = headers.indexOf("Rank");

  const districtRows = rows.filter(row =>
    String(row[districtCol]).toLowerCase().trim() === districtName.toLowerCase().trim()
  );

  const crashTypes = [...new Set(districtRows.map(row => String(row[crashTypeCol]).trim()))];

  let insertIndex = presentation.getSlides().indexOf(templateSlide);

  for (const type of crashTypes) {
    const typeRows = districtRows
      .filter(row => String(row[crashTypeCol]).trim() === type)
      .sort((a, b) => Number(b[fatalitiesCol]) - Number(a[fatalitiesCol]));

    if (typeRows.length === 0) continue;

    // Re-assign ranks based on fatalities
    for (let i = 0; i < typeRows.length; i++) {
      typeRows[i][rankColIndex] = i + 1;
    }

    const chunkSize = 12;
    const totalFatalities = typeRows.reduce((sum, row) => sum + Number(row[fatalitiesCol]) || 0, 0);
    const totalAllFatalities = districtRows.reduce((sum, row) => sum + Number(row[fatalitiesCol]) || 0, 0);
    const percentage = totalAllFatalities === 0 ? 0 : ((totalFatalities / totalAllFatalities) * 100).toFixed(1);

    const headingText = `Critical Locations: ${type} (2023 & 2024)`;
    const countText = `${typeRows.length} Critical Locations account for ~${percentage}% fatalities due to ${type}`;

    for (let i = 0; i < typeRows.length; i += chunkSize) {
      const chunk = typeRows.slice(i, i + chunkSize);
      const duplicatedSlide = templateSlide.duplicate();
      insertIndex = presentation.getSlides().indexOf(duplicatedSlide); // Reset insertIndex

      // Update title and count line
      for (const shape of duplicatedSlide.getShapes()) {
        if (shape.getText) {
          const txt = shape.getText().asString().trim().toLowerCase();
          if (txt.includes("critical locations:")) {
            shape.getText().setText(headingText);
          } else if (txt.includes("critical locations account for")) {
            shape.getText().setText(countText);
          }
        }
      }

      const table = duplicatedSlide.insertTable(chunk.length + 1, selectedColumns.length);

    // Header row
    selectedColumns.forEach((header, col) => {
      const cellText = table.getCell(0, col).getText();
      cellText.setText(header);
      cellText.getTextStyle()
        .setBold(true)
        .setFontFamily("Montserrat")
        .setFontSize(8);
    });

    // Data rows
    chunk.forEach((row, rowIndex) => {
      columnIndexes.forEach((colIndex, col) => {
        const cellText = table.getCell(rowIndex + 1, col).getText();
        cellText.setText(String(row[colIndex] || ""));
        cellText.getTextStyle()
          .setFontFamily("Montserrat")
          .setFontSize(8);
      });
    });


      Logger.log(`Inserted ${chunk.length} rows for crash type "${type}"`);
    }
  }
}
