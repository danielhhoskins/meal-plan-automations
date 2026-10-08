/**
 * The event handler triggered when editing the spreadsheet.
 * @param {Event} e The onEdit event.
 * @see https://developers.google.com/apps-script/guides/triggers#onedite
 */

function onEdit(e) {
  changeMacrosBasedOnChangeToAmount(e);
  newValueOfIndividualMacro(e);
  addAmountToIndividualMacro(e);
  searchForFoodAndFillRow(e);
  clearRowIfFoodCellIsDeleted(e);
  removeBlankRows(e);
}

function indexOfMax(arr) {
    if (arr.length === 0) {
        return -1;
    }

    var max = arr[0];
    var maxIndex = 0;

    for (var i = 1; i < arr.length; i++) {
        if (arr[i] > max) {
            maxIndex = i;
            max = arr[i];
        }
    }
    return maxIndex;
}

function changeMacrosBasedOnChangeToAmount(e) {
  Logger.log("Initial Log");
  let editedCol = e.range.getColumn();
  if (editedCol != 4) return;
  const range = e.range;
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheets()[0];
  let editedRow = e.range.getRow();
  Logger.log("editedRow: " + editedRow);
  let pRange = sheet.getRange(editedRow, 5);
  let cRange = sheet.getRange(editedRow, 6);
  let fRange = sheet.getRange(editedRow, 7);
  Logger.log("second log");
  let pVal = parseFloat(String(pRange.getValue()));
  Logger.log("pVal: " + pVal);
  let cVal = parseFloat(String(cRange.getValue()));
  Logger.log("cVal: " + cVal);
  let fVal = parseFloat(String(fRange.getValue()));
  Logger.log("fVal: " + fVal);
  if (isNaN(pVal) || isNaN(cVal) || isNaN(fVal)) return; // Make sure all macro values are floats
  let oldVal = String(e.oldValue);
  Logger.log("oldVal: " + oldVal);
  if (oldVal.includes("/")) return;
  let oldNum = parseFloat(oldVal);
  Logger.log('oldNum: ' + oldNum);
  Logger.log("about to enter getNumber");
  let newNum = parseFloat(e.value);
  Logger.log('newNum: ' + newNum);
  let oldRemainingChars = oldVal.match(/(?<=\d+)\D*$/);
  Logger.log('oldRemainingChars: ' + oldRemainingChars);
  let newRemainingChars = /(?<=\d+)\D+$/.test(e.value) ? e.value.match(/(?<=\d+)\D*$/) : oldRemainingChars; // Are there characters after the numerical chars? i.e. unit characters
  // let newRemainingChars = e.value.match(/(?<=\d+)\D*$/);
  Logger.log('newRemainingChars: ' + newRemainingChars);
  let comp = String(newRemainingChars).trim().toLowerCase().replace(/s$/, '') === String(oldRemainingChars).trim().toLowerCase().replace(/s$/, '');
  Logger.log("comp: " + comp)
  if (!comp) return;
  if (isNaN(oldNum)) return;
  if (isNaN(newNum)) return;
  // Format of column D must be "Plain Text"
  let integerVal = oldVal.match(/^\d+(\s|[a-zA-Z])/); // Integer match
  Logger.log("integerVal: " + integerVal);
  let decimalVal = oldVal.match(/^\d+\.\d+/); // Decimal match
  Logger.log("decimalVal: " + decimalVal);
  let fractionVal = oldVal.match(/^\d+\/\d+/); // Fraction match
  Logger.log("fractionVal: " + fractionVal);
  let multiplier = newNum / oldNum;
  Logger.log('multiplier: ' + multiplier);
  let newPVal = parseFloat(pVal * multiplier);
  Logger.log("newPVal: " + newPVal);
  // range.setNote("Ran");
  let newCVal = parseFloat(cVal * multiplier);
  Logger.log("newCVal: " + newCVal);
  let newFVal = parseFloat(fVal * multiplier);
  Logger.log("newFVal: " + newFVal);
  pRange.setValue(newPVal);
  cRange.setValue(newCVal);
  fRange.setValue(newFVal);
  let quantityCellRange = sheet.getRange(editedRow, 4);
  quantityCellRange.setValue(newNum + newRemainingChars);
}

function newValueOfIndividualMacro(e) {
  // If you enter "<number><p c or g>" in column I (column9), it recalculates the other two macros
  let editedCol = e.range.getColumn();
  if (editedCol != 9) return;
  let newVal = e.value;
  let lastChar = e.value.slice(-1);
  if (lastChar === "+" || lastChar === "-") return;
  if (lastChar !== "p" && lastChar !== "c" && lastChar !== "f") return;
  let newNum = parseFloat(e.value);
  if (Array.from(e.value)[0] === "+" || Array.from(e.value)[0] === "-") return;
  if (isNaN(newNum)) {
    Logger.log("Doesn't begin with number.");
    return;
  }
  const range = e.range;
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheets()[0]
  let editedRow = e.range.getRow();
  Logger.log(sheet.getRange(editedRow, 1).getValue());
  if (sheet.getRange(editedRow, 1).getValue() === "Meal Total") return;
  Logger.log("editedRow: " + editedRow);
  let pRange = sheet.getRange(editedRow, 5);
  let cRange = sheet.getRange(editedRow, 6);
  let fRange = sheet.getRange(editedRow, 7);
  let pVal = parseFloat(String(pRange.getValue()));
  let cVal = parseFloat(String(cRange.getValue()));
  let fVal = parseFloat(String(fRange.getValue()));
  if (isNaN(pVal) || isNaN(cVal) || isNaN(fVal)) return; // Make sure all macro values are floats
  let multiplier;
  if (lastChar === "p") {multiplier = newNum / pVal;
  } else if (lastChar === "c") {multiplier = newNum / cVal;
  } else if (lastChar === "f") {multiplier = newNum / fVal;}
  let newPVal = parseFloat(pVal * multiplier);
  let newCVal = parseFloat(cVal * multiplier);
  let newFVal = parseFloat(fVal * multiplier);
  let quantityCellRange = sheet.getRange(editedRow, 4);
  let quantityCell = String(quantityCellRange.getValue());
  if (quantityCell.includes("/")) return;
  pRange.setValue(newPVal);
  cRange.setValue(newCVal);
  fRange.setValue(newFVal);
  let oldQuantityNum = parseFloat(quantityCell);
  let unitChars = quantityCell.match(/(?<=\d+)\D*$/);
  let newQuantityNum = parseFloat((oldQuantityNum * multiplier).toFixed(2));
  Logger.log('newQuantityNum: ' + newQuantityNum);
  quantityCellRange.setValue(newQuantityNum + unitChars);
  range.clearContent();
}

function addAmountToIndividualMacro(e) {
  if (isNaN(parseFloat(e.value))) return;
  let newNum = parseFloat(e.value);
  let lastChar = e.value.slice(-1);
  if (lastChar !== "+" && lastChar !== "-") return;
  const macroLetter = e.value.slice(-2,-1);
  const range = e.range;
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheets()[0]
  let editedRow = e.range.getRow();
  if (sheet.getRange(editedRow, 1).getValue() === "Meal Total") return;
  let pRange = sheet.getRange(editedRow, 5);
  let cRange = sheet.getRange(editedRow, 6);
  let fRange = sheet.getRange(editedRow, 7);
  let pVal = parseFloat(String(pRange.getValue()));
  let cVal = parseFloat(String(cRange.getValue()));
  let fVal = parseFloat(String(fRange.getValue()));
  if (lastChar === "+") {
    if (macroLetter === "p") {multiplier = (pVal + newNum) / pVal;
    } else if (macroLetter === "c") {multiplier = (cVal + newNum) / cVal;
    } else if (macroLetter === "f") {multiplier = (fVal + newNum) / fVal;}
  } else if (lastChar === "-") {
    if (macroLetter === "p") {multiplier = (pVal - newNum) / pVal;
    } else if (macroLetter === "c") {multiplier = (cVal - newNum) / cVal;
    } else if (macroLetter === "f") {multiplier = (fVal - newNum) / fVal;}
  }
  let newPVal = parseFloat(pVal * multiplier);
  let newCVal = parseFloat(cVal * multiplier);
  let newFVal = parseFloat(fVal * multiplier);
  let quantityCellRange = sheet.getRange(editedRow, 4);
  let quantityCell = String(quantityCellRange.getValue());
  if (quantityCell.includes("/")) return;
  pRange.setValue(newPVal);
  cRange.setValue(newCVal);
  fRange.setValue(newFVal);
  let oldQuantityNum = parseFloat(quantityCell);
  let unitChars = quantityCell.match(/(?<=\d+)\D*$/);
  let newQuantityNum = parseFloat((oldQuantityNum * multiplier).toFixed(2));
  Logger.log('newQuantityNum: ' + newQuantityNum);
  quantityCellRange.setValue(newQuantityNum + unitChars);
  range.clearContent();
}

function replacePunctuationWithWhiteSpace(str) {
  str = str.replace(/[^\w\s\']|_/g, "").replace(/\s+/g, " ");
  return str;
}

function isMatch(enteredStr, str2) {
  enteredStr = enteredStr.toLowerCase();
  str2 = str2.toLowerCase();
  if(str2.includes(enteredStr)) {
    Logger.log("Found match of enteredStr: " + enteredStr + " and str2: " + str2);
    return true;
  }
  else return false;
}

function copyData(ss, sourceRange, destinationRangeStart){
  Logger.log("copying data from sourceRange " + sourceRange + " to destinationRangeStart" + destinationRangeStart);
  // Gather the source range values
  const sourceRng = ss.getRange(sourceRange)
  const sourceVals = sourceRng.getValues();
 
  const destStartRange = ss.getRange(destinationRangeStart);
  const destSheet = destStartRange.getSheet();
 
  // Get the full data range to paste from start range.
  const destRange = destSheet.getRange(
      destStartRange.getRow(),
      destStartRange.getColumn(),
      sourceVals.length,
      sourceVals[0].length
    );
  
  // Paste in the values.
  destRange.setValues(sourceVals);
  // SpreadsheetApp.flush();
};

function searchForFoodAndFillRow(e) {
  function convertTokensToLowercase(tokensArr) {
    let lcToks = new Array();
    for (ucTok of tokensArr) lcToks.push(ucTok.toLowerCase());
    return lcToks;
  }
  function getBestMatch(enteredStr, e) {
    function containsAllTokens(potentialResult, enteredStrToks) { // Verify that the potentialResult contains all tokens of enteredStr
      let prToks = potentialResult.split(' ');
      prToks = convertTokensToLowercase(prToks);
      // let newPRToks = new Array();
      // for (prtok of prToks) {
      //   newPRToks.push(prtok.toLowerCase());
      // }
      // prToks = newPRToks;
      for (tok of enteredStrToks) {
        if (!prToks.includes(tok)) {
          // Logger.log("potentialResult doesn't contain all tokens of enteredStr: " + tok)
          return false;
        }
      }
      return true;
    }
    Logger.log("enteredStr: " + enteredStr);
    let enteredStrToks = enteredStr.split(' ');
    enteredStrToks = convertTokensToLowercase(enteredStrToks);
    Logger.log("enteredStrToks: " + enteredStrToks);
    let potentialResults = new Array();
    // Start with array with each element being a subarray with the food title and its associated row number
    for (i=0;i<data.length;i++){ // Loop through potential results (food rows) and create new potential results array that doesn’t include any that don’t include all tokens in the query.
      // Logger.log("typeof data[i][0]: " + typeof data[i][0]);
      // Logger.log("data[i][0]: " + data[i][0]);
      // Logger.log("potentialResults: " + potentialResults);
      // Logger.log("potentialResults[1]: " + potentialResults[1]);
      // Logger.log("containsAllTokens(data[i][0], enteredStr): " + containsAllTokens(data[i][0], enteredStr));
      if (containsAllTokens(data[i][0], enteredStrToks)) {
        potentialResults.push([data[i][0], i]);
      }
    }
    Logger.log("potentialResults: " + potentialResults);
    // Calculate the percentage of each potential result’s tokens included in the query.
    let percentages = new Array();
    for (const pr of potentialResults) {
      let prStr = pr[0];
      let prToks = convertTokensToLowercase(prStr.split(' '));
      let cnt = 0;
      for (prtok of prToks) {
        if (enteredStrToks.includes(prtok)) cnt++;
      }
      percentages.push(cnt / prToks.length);
    }
    Logger.log("percentages: " + percentages);
    Logger.log("Math.max( ...percentages ): " + Math.max( ...percentages ));
    if (Math.max( ...percentages ) > 0) { // If there’s at least one with a percentage higher than 0: Return the row with the highest percentage
      let idxOfMax = indexOfMax(percentages);
      Logger.log("idxOfMax: " + idxOfMax);
      let rowIndex = potentialResults[idxOfMax][1];
      Logger.log("rowIndex: " + rowIndex);
      return rowIndex;
    } else { // If there are none with a percentage higher than 0: Run the existing algorithm which matches exact string
      for (i=0;i<data.length;i++){ // Check each row for a match
        // Logger.log("data[i][0]: " + data[i][0]);
        if (isMatch(enteredStr, data[i][0])) {
          Logger.log("getBestMatch returning i: " + i);
          return i;
          // Logger.log(`Found a match between enteredStr ${enteredStr} and data[i][0] ${data[i][0]}`);
        }
      }
    }
  }
  if (e.range.getColumn() != 2) return;
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let enteredStr = e.range.getValue().toLowerCase();
  if (enteredStr === "") return;
  enteredStr = replacePunctuationWithWhiteSpace(enteredStr).trim();
  let mealPlanSheet = ss.getSheets()[0];
  // let foodsSheet = ss.getSheets()[1];
  let editedRow = e.range.getRow();
  Logger.log("editedRow: " + editedRow);
  Logger.log("Checking columns C through G")
  if (!ss.getRange("Meal Plan!C" + editedRow + ":G" + editedRow).isBlank()) {
    Logger.log("Row " + editedRow + " is not empty.");
    return;
  }
  let data = ss.getSheetByName("Foods").getDataRange().getValues();
  let matchedRowIndex = getBestMatch(enteredStr, e);
  copyData(e.source, `Foods!A${matchedRowIndex+1}:F${matchedRowIndex+1}`, `Meal Plan!B${editedRow}:G${editedRow}`);
  Logger.log("Done moving macros")
}

function clearRow(rowNumberOneIndexed) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheets()[0];
  sheet.getRange(rowNumberOneIndexed, 2, 1, 6).clearContent();
}

function clearRowIfFoodCellIsDeleted(e) {
  let editedCol = e.range.getColumn();
  if (editedCol != 2) return;
  Logger.log("e.range.getNumColumns(): " + e.range.getNumColumns());
  if (e.range.getNumColumns() != 1) return; 
  let enteredStr = e.range.getValue().toLowerCase();
  let firstRow = e.range.getRow();
  if (enteredStr === "") {
    Logger.log("attempting to clear row: " + firstRow);
    for (i=firstRow; i<firstRow+e.range.getNumRows(); i++){
      Logger.log("Clearing row " + i);
      clearRow(i);
    }
  }
}


function removeBlankRows(e) {
  // function consolidateRowsInMeal(firstRow, lastRow) { // oldest
  //   for (i=firstRow;i<=lastRow+1;i++){
  //     Logger.log("i loop iter. i: " + i);
  //     let nextNonBlankrow;
  //     if (ss.getRange("Meal Plan!B" + i + ":G" + i).isBlank()) {
  //       nextNonBlankrow = -1;
  //       for (j=i;j<=lastRow+1;j++) { // find the next non-blank row
  //         Logger.log("j loop iter. j: " + j);
  //         if (!ss.getRange("Meal Plan!B" + j + ":G" + j).isBlank()) {
  //           Logger.log("Found next non-blank row: " + j);
  //           nextNonBlankrow = j;
  //           copyData(e.source, `Meal Plan!B${nextNonBlankrow}:G${nextNonBlankrow}`, `Meal Plan!B${i}:G${i}`);
  //           clearRow(nextNonBlankrow);
  //           break;
  //         }
  //         if (j == lastRow+1) {
  //           console.log("No more nonblank rows in this section.")
  //           i = lastRow+1; // breack the outer for loop (i.e. the i for loop)
  //         }
  //       }
  //     }
  //   }
  // }
  // function consolidateRowsInMeal(firstRow, lastRow) { // second oldest
  //   for (i=firstRow;i<=lastRow+1;i++){
  //     Logger.log("i loop iter. i: " + i);
  //     let nextNonBlankrow;
  //     Logger.log("rowIsBlank(data[i-1]): " + rowIsBlank(data[i-1]));
  //     if (rowIsBlank(data[i-1])) {
  //       nextNonBlankrow = -1;
  //       for (j=i;j<=lastRow+1;j++) { // find the next non-blank row
  //         Logger.log("j loop iter. j: " + j);
  //         Logger.log("rowIsBlank(data[j-1]): " + rowIsBlank(data[j-1]));
  //         if (!rowIsBlank(data[j-1])) {
  //           Logger.log("Found next non-blank row: " + j);
  //           nextNonBlankrow = j;
  //           copyData(e.source, `Meal Plan!B${nextNonBlankrow}:G${nextNonBlankrow}`, `Meal Plan!B${i}:G${i}`);
  //           clearRow(nextNonBlankrow);
  //           break;
  //         }
  //         if (j == lastRow+1) {
  //           console.log("No more nonblank rows in this section.")
  //           i = lastRow+1; // breack the outer for loop (i.e. the i for loop)
  //         }
  //       }
  //     }
  //   }
  // }
  function consolidateRowsInMeal(firstRow, lastRow) {
    Logger.log("firstRow: " + firstRow);
    Logger.log("lastRow: " + lastRow);
    let newData = new Array();
    for (i=firstRow; i<=lastRow; i++) { // these row indexes match those in the actual spreaadsheet. Aka not zero indexed
      let newRow = new Array(6);
      if (!rowIsBlank(i, data)) {
        newData.push(data[i].slice(1,7));
      }
    }
    let numBlankRowsToAdd = (lastRow-firstRow+1) - newData.length;
    for (i=0; i<numBlankRowsToAdd; i++) {
      newData.push(['','','','','','']);
    }
    for (i=0; i<newData.length; i++) {
      Logger.log("i: " + i);
      Logger.log("newData[i]: " + newData[i]);
    }
    sheet.getRange(firstRow+1, 2, lastRow-firstRow+1, 6).setValues(newData);
  }
  function rowIsBlank(rowIdx, data) {
    // Check whether a row is blank (except for calories column)
    let row = data[rowIdx];
    for (n=1; n<=6; n++) {
      // Logger.log("row[" + n + "]: " + row[n]);
      if (row[n] !== "") return false;
    }
    return true;
  }
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Meal Plan");
  let data = sheet.getDataRange().getValues();
  // Logger.log("data: " + data);
  // for (row of data) {
  //   Logger.log("rowIsBlank(row): " + rowIsBlank(row));
  //   // Logger.log("Object.prototype.toString.call(row): " + Object.prototype.toString.call(row));
  // }
  // for (let [index, row] of data.entries()) {
   // your code goes here    
    // Logger.log("index: " + index);
    // Logger.log("rowIsBlank(row): " + rowIsBlank(row));
  // }
  let breakfastRow;
  let lunchRow;
  let afternoonSnackRow;
  let dinnerRow;
  let bedtimeRow;
  let totalRow;
  for (i=0;i<data.length;i++){
    if (data[i][0].includes("Breakfast")) breakfastRow = i;
    if (data[i][0].includes("Lunch")) lunchRow = i;
    if (data[i][0].includes("Dinner")) dinnerRow = i;
    if (data[i][0].includes("Afternoon")) afternoonSnackRow = i;
    if (data[i][0].includes("Bedtime")) bedtimeRow = i;
    if (data[i][0].includes("Target")) totalRow = i-1;
  }
  Logger.log("blunch");
  consolidateRowsInMeal(breakfastRow, lunchRow-2);
  Logger.log("las");
  consolidateRowsInMeal(lunchRow, afternoonSnackRow-2);
  Logger.log("asdinner");
  consolidateRowsInMeal(afternoonSnackRow, dinnerRow-2);
  Logger.log("dedtime");
  consolidateRowsInMeal(dinnerRow, bedtimeRow-2);
  Logger.log("bedtimetoend")
  consolidateRowsInMeal(bedtimeRow, totalRow-2);
  
}