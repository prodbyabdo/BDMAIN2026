function convertBgetnpiToValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  // Regex to exactly target the =bgetnpi custom function (case-insensitive)
  var targetRegex = /=bgetnpi/i;

  // Loop through all 19 tabs
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var range = sheet.getDataRange();
    
    var formulas = range.getFormulas();
    var values = range.getValues();
    var formulasChanged = false;

    // Scan every row and column in the tab
    for (var row = 0; row < formulas.length; row++) {
      for (var col = 0; col < formulas[row].length; col++) {
        var currentFormula = formulas[row][col];
        
        // If the formula contains bgetnpi, swap it out for its static value
        if (currentFormula !== "" && targetRegex.test(currentFormula)) {
          formulas[row][col] = values[row][col];
          formulasChanged = true;
        }
      }
    }

    // Only update the grid if targeted formulas were found and stripped
    if (formulasChanged) {
      range.setValues(formulas);
    }
  }
}
