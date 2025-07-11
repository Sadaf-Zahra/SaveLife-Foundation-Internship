    function extractTextFromSlideDynamic() {
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();
  var targetHeading = "Summary Recommendations - Mandya";
  var text = '';

  for (var i = 0; i < slides.length; i++) {
    var slide = slides[i];
    var shapes = slide.getShapes();
    for (var j = 0; j < shapes.length; j++) {
      var shape = shapes[j];
      if (shape.getText && shape.getText().asString().includes(targetHeading)) {
        processShapesRecursively(slide.getShapes(), text);
        if (text.trim() === '') {
          Logger.log("No text extracted from shapes.");
        } else {
          Logger.log(text);
        }
        return text;
      }
    }
  }

  Logger.log("Heading not found.");
  return "Heading not found.";
}

function processShapesRecursively(shapes, text) {
  for (var i = 0; i < shapes.length; i++) {
    var shape = shapes[i];
    if (shape.getText) {
      var shapeText = shape.getText().asString();
      if (shapeText) text += shapeText + '\n';
    } else if (shape.getTable) {
      var table = shape.getTable();
      for (var row = 0; row < table.getNumRows(); row++) {
        for (var cell = 0; cell < table.getRow(row).getNumCells(); cell++) {
          var cellText = table.getCell(row, cell).getText();
          if (cellText) text += cellText.asString() + '\n';
        }
      }
    } else if (shape.getGroup) {
      processShapesRecursively(shape.getGroup().getChildren(), text);
    } else {
      Logger.log("Unhandled shape type: " + shape.getShapeType());
    }
  }
}