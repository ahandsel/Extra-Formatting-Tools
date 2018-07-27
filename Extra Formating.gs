// Extra Formatting {Google Sheets}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

// Scripts Creats a "Custom menu" w/ "Capitalize Each Word", "lower case", & "UPPER CASE" buttons
function onOpen() {
	// Adds the Custom menu to the Active Spreadsheet
	SpreadsheetApp.getUi()
		.createMenu('Extra Formating')
			.addItem('Capitalize Each Word', 'proper')
			.addItem('lower case', 'lower_cap')
			.addItem('UPPER CASE', 'upper_cap')
			.addItem('Delete Empty Rows', 'deleteEmptyRows')
			.addItem('Even Row to Next Column', 'evenRow_nextColumn')
			.addSeparator()
			.addToUi();
}

// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

// Interface between the Custom Menu & Main function for text formating
function proper() {	txt_format_main(proper); }
function lower_cap() {	txt_format_main(lower_cap); }
function upper_cap() {	txt_format_main(upper_cap); }

// = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

// Main function for text formating {Capitalize Each Word, lower case, & UPPER CASE}
function txt_format_main(format_type) {
	var tempArray = []; // Temp Array that holds the text for convertion
	var sheet = SpreadsheetApp.getActiveSheet();
	var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	
	// Determine the selected cells & returns their location in A1 notation {e.g. B8:D10}
	var selectedAddress = spreadsheet.getActiveSheet().getSelection().getActiveRange().getA1Notation();
	Index_Array = A1_to_Index(selectedAddress); // Index_Array = Row 1, Column 1, Row 2, Column 2
	var range = sheet.getRange(Index_Array[0], Index_Array[1], Index_Array[2], Index_Array[3]); 
		//getRange(row, column, numRows, numColumns)
	
	// Extracting data from the Selected Cells
	var input_values = range.getValues(); // input_values now holds the strings that we want to convert

	// Messages for testing purpose
	Browser.msgBox("Input Values = " + input_values);
	
	// Iterate through an array, change the string formating, & insert the updated string back into the selected cells
	// **** THIS IS WHERE THE PROGRAM FAILS! ***
	for (var row in input_values) {
		for (var col in input_values[row]) {
			
			Browser.msgBox( input_values[row][col] ); //<= input is now working {for testing purpose}
			tempArray.push( [ toTitleCase( input_values[row][col] ) ] ); //<=  working-ish
			Browser.msgBox(tempArray); //<= Convertion is now working {for testing purpose}
			
			sheet.getRange(input_values[row][col]).setValue( [ toTitleCase( input_values[row][col] ) ] );
			
			//range.setValues( toTitleCase(values[row][col]) );
			Browser.msgBox( input_values[row][col] ); //<= input is now working

		}
	}
	
/*
	var temp_Index_Array = Index_Array; // [0] = Row 1 | [1] = Column 1
	
	//var temp_range = sheet.getRange(Index_Array[0], Index_Array[1]); //getRange(row, column) at Start
	var temp_value = sheet.getRange(temp_Index_Array[0], temp_Index_Array[1]).getValue();
	
	Browser.msgBox( "temp_Index_Array = " + temp_Index_Array) // for testing purpose
	Browser.msgBox( "numRows = " + Index_Array[2]) // for testing purpose
	Browser.msgBox( "numColumns = " + Index_Array[3]) // for testing purpose
	
	for (var r=1; r<Index_Array[2]; r++) { 
		for (var c=1; c<Index_Array[3]; c++) { 
			
			Browser.msgBox( 'r = '+r); // for testing purpose
			Browser.msgBox( 'c = '+c); // for testing purpose
			
			Browser.msgBox( 'Temp Row = '+temp_Index_Array[0]); // for testing purpose
			Browser.msgBox( 'Temp Col = '+temp_Index_Array[1]); // for testing purpose

			sheet.getRange(temp_Index_Array[0], temp_Index_Array[1])
				.setValue( [ toTitleCase( temp_value ) ] );
			
			++temp_Index_Array[1]; // ++Column
			temp_value = sheet.getRange(temp_Index_Array[0], temp_Index_Array[1]).getValue();
			
		}
		++temp_Index_Array[0]; // ++Row
	}
	*/
	
/*
// Working-ish code
	if (format_type == proper) {
		range.getValues().forEach(
		//sheet.getRange(selectedAddress).getValues().forEach(
			function (r) { tempArray.push([toTitleCase(r[0])]) }
		);
	} else if (format_type == lower_cap){
		sheet.getRange(selectedAddress).getValues().forEach(
			function (r) { tempArray.push([lowercase(r[0])]) }
		);
	} else if (format_type == upper_cap){
		sheet.getRange(selectedAddress).getValues().forEach(
			function (r) { tempArray.push([uppercase(r[0])]) }
		);
	}
	
	Browser.msgBox("tempArray " + tempArray); //for testing purpose

	sheet.getRange(selectedAddress).setValues(tempArray);
	*/
}

// = = = = = = = = = = = = = = = = = = = = = = = = = = = =

function toTitleCase(str) {
	return str.replace(/\w\S*/g, function (txt) {
		return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
	});
}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = =
function uppercase(str) {
	return str.replace(/\w\S*/g, function (txt) {
		return txt.substr(0).toUpperCase();
	});
}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = =
function lowercase(str) {
	return str.replace(/\w\S*/g, function (txt) {
		return txt.substr(0).toLowerCase();
	});
}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = =

// Converts A1 Notation to Array with Row 1, Col 1, Row 2, Col 2 Index
// for getRange(row, column, numRows, numColumns)
function A1_to_Index(A1Notation) {
	
	// Converting A1 Notation to Column & Row Index
	Address_Array = A1Notation.split(":");			//e.g. B8:D10 => B8, D10 as Array

	col_1_Index = Letter_to_Num(Address_Array[0].charAt(0));		//e.g. B => 2
	col_2_Num = Letter_to_Num(Address_Array[1].charAt(0))- col_1_Index + 1;		//e.g. D => 4

	row_1_Index = Address_Array[0].substring(1);	//e.g. 8
	row_2_Num = Address_Array[1].substring(1) - row_1_Index + 1;	//e.g. 10-8 = 2
	
	// Fixing for A:A issues {Column to Column w/out 
	if (row_1_Index == '') { row_1_Index = 1 }
	if (row_2_Num < 1) { row_2_Num = 1 }

	// Index_Array = Row 1, Column 1, Row 2, Column 2
	var Index_Array = [row_1_Index, col_1_Index, row_2_Num, col_2_Num];
	
	return Index_Array;
}

function Letter_to_Num(col_Letter) {
	// Convert the A1 Notation's Column Letter to a Number
	// A = 1, BB = 27, etc.
	var alphabet = ["A","B","C","D","E","F","G","H","I","J","K","L","M", "N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","BB","CC","DD","EE","FF","GG","HH","II","JJ","KK","LL","MM", "NN","OO","PP","QQ","RR","SS","TT","UU","VV","WW","XX","YY","ZZ"];
	var col_Index = alphabet.indexOf(col_Letter)+1;
	
	// Messages for testing purpose
	//Browser.msgBox("Is " + col_Letter + ' = ' +col_Index);
	
	return col_Index;
}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// = = = = = = = = = = = = = = = = = = = = = = = = = = = =

// = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// = = = = = = = = = = = = = = = = = = = = = = = = = = = =

function deleteEmptyRows(){ 
	var sheet = SpreadsheetApp.getActiveSheet();
	var data = sheet.getDataRange().getValues();
	var targetData = new Array();
	for( n=0; n<data.length; ++n){
		if(data[n].join().replace(/,/g,'')!=''){	targetData.push(data[n])	};
		Logger.log(data[n].join().replace(/,/g,''))
	}
	sheet.getDataRange().clear();
	sheet.getRange(1,1,targetData.length,targetData[0].length).setValues(targetData);
}

// = = = = = = = = = = = = = = = = = = = = = = = = = = = =
// = = = = = = = = = = = = = = = = = = = = = = = = = = = =

function evenRow_nextColumn() {
	var sheet = SpreadsheetApp.getActiveSheet()
			
	// This represents ALL the data
	var range = sheet.getDataRange();
	var values = range.getValues();
	
	var targetData = new Array();
	for (var column = 1; column <= range.getNumColumns(); column++) {
		for (var row = 1; row <= range.getNumRows(); row++) {
			if (row % 2 == 0){ 
				sheet.getRange("A1:E").moveTo(sheet.getRange("F1"));
			}
		}
	}
}
// = = = = = = = = = = = = = = = = = = = = = = = = = = = =
