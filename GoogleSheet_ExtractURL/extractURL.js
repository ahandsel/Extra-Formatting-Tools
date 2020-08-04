// Returns the URL of a hyperlinked cell

function extractURL(reference) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var args = formula.match(/=\w+\((.*)\)/i);
  try {
    var range = sheet.getRange(args[1]);
  } catch (e) {
    throw new Error(args[1] + ' is not a valid range');
  }
  var formulas = range.getFormulas();
  var output = [];
  for (var i = 0; i < formulas.length; i++) {
    var row = [];
    for (var j = 0; j < formulas[0].length; j++) {
      var url = formulas[i][j].match(/=hyperlink\("([^"]+)"/i);
      row.push(url ? url[1] : '');
    }
    output.push(row);
  }
  return output;
}

function extractURL2(input) {
  var range = SpreadsheetApp.getActiveSheet().getRange(input);
  var url = /"(.*?)"/.exec(range.getFormulaR1C1())[1];
  return url;
}