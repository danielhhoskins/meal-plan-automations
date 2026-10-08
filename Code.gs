/**
 * The event handler triggered when editing the spreadsheet.
 * @param {Event} e The onEdit event.
 * @see https://developers.google.com/apps-script/guides/triggers#onedite
 */

// Runs from an installable trigger (see installEditTrigger) because simple onEdit
// triggers aren't allowed to call UrlFetchApp, which the LLM food lookup needs.
function onEditInstalled(e) {
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
  const choice = chooseFoodsWithLlm(String(e.range.getValue()), data);
  if (choice === null) { // LLM call failed: use the old algorithm
    let matchedRowIndex = getBestMatch(enteredStr, e);
    copyData(e.source, `Foods!A${matchedRowIndex+1}:F${matchedRowIndex+1}`, `Meal Plan!B${editedRow}:G${editedRow}`);
    Logger.log("Done moving macros")
    return;
  }
  if (choice.fillRemaining) {
    fillRemainingCalories(ss, editedRow, data);
    return;
  }
  const items = choice.items;
  if (items.length === 0) {
    Logger.log("LLM found no food in: " + enteredStr);
    return;
  }
  const foodsSheet = ss.getSheetByName("Foods");
  let nextFoodsRow = data.length;
  while (nextFoodsRow > 1 && String(data[nextFoodsRow - 1][0]).trim() === "") nextFoodsRow--;
  nextFoodsRow++;
  const mealRows = [];
  for (const item of items) {
    let foodRow, amount;
    if (item.row !== -1) {
      foodRow = data[item.row].slice(0, 6);
      amount = item.amount;
    } else { // Not in Foods: look it up online and add it
      const newFood = lookUpNewFoodWithLlm(item.text);
      if (newFood === null) {
        Logger.log("No food added for: " + item.text);
        continue;
      }
      foodRow = [newFood.food, newFood.brand, newFood.serving, newFood.protein, newFood.carbs, newFood.fat];
      foodsSheet.getRange(nextFoodsRow, 1, 1, 7)
        .setValues([foodRow.concat([`=D${nextFoodsRow}*4+E${nextFoodsRow}*4+F${nextFoodsRow}*9`])]);
      Logger.log("Added " + newFood.food + " to Foods row " + nextFoodsRow);
      nextFoodsRow++;
      amount = newFood.amount;
    }
    const scaledRow = amount !== null ? scaleFoodRow(foodRow, amount) : null; // Scaled before writing so the Foods amount never shows
    mealRows.push(scaledRow !== null ? scaledRow : foodRow);
  }
  writeRowsIntoMeal(ss.getSheetByName("Meal Plan"), editedRow, mealRows);

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

const LLM_MODEL = "claude-opus-5-5";

// Sends a Messages API request and returns the JSON object in Claude's final text block, or null if the
// call failed. Continues the request if a web search pauses the turn (stop_reason "pause_turn").
function callClaudeForJson(body) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");
  if (!apiKey) {
    Logger.log("No ANTHROPIC_API_KEY script property set");
    return null;
  }
  const messages = body.messages.slice();
  try {
    for (let attempt = 0; attempt < 5; attempt++) {
      const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
        method: "post",
        contentType: "application/json",
        headers: {
          "x-api-key": apiKey,
          "anthropic-version": "2023-06-01",
          "anthropic-beta": "server-side-fallback-2026-07-01"
        },
        payload: JSON.stringify(Object.assign({}, body, { messages: messages })),
        muteHttpExceptions: true
      });
      if (response.getResponseCode() !== 200) {
        Logger.log("LLM call failed: " + response.getResponseCode() + " " + response.getContentText());
        return null;
      }
      const message = JSON.parse(response.getContentText());
      if (message.stop_reason === "pause_turn") {
        messages.push({ role: "assistant", content: message.content });
        continue;
      }
      if (message.stop_reason === "refusal") {
        Logger.log("LLM refused: " + JSON.stringify(message.stop_details));
        return null;
      }
      const textBlocks = message.content.filter(b => b.type === "text");
      return JSON.parse(textBlocks[textBlocks.length - 1].text);
    }
    Logger.log("LLM call kept pausing; giving up");
    return null;
  } catch (err) {
    Logger.log("LLM call error: " + err);
    return null;
  }
}

// Asks Claude which foods the entered text describes. A single entry can describe several foods (e.g. a
// sandwich), so this returns a list of {row, text, amount}: row is the 0-indexed row in data (-1 if the food
// isn't in Foods), text describes that one food for an online lookup, and amount is its quantity in the unit
// of that row's Amount (null to use the Foods amount). Returns {fillRemaining, items}: items is [] if the text names no
// food, and fillRemaining is true if the text asks to fill the day's remaining calories. Returns null if the call failed.
function chooseFoodsWithLlm(enteredText, data) {
  let foodList = "";
  for (let i = 1; i < data.length; i++) { // Row 0 is the header
    if (String(data[i][0]).trim() === "") continue;
    foodList += i + ": " + data[i][0] + (data[i][1] ? " (" + data[i][1] + ")" : "") + " | " + data[i][2] + "\n";
  }
  const result = callClaudeForJson({
    model: LLM_MODEL,
    max_tokens: 8000,
    fallbacks: "default",
    output_config: {
      effort: "medium",
      format: {
        type: "json_schema",
        schema: {
          type: "object",
          properties: {
            items: {
              type: "array",
              items: {
                type: "object",
                properties: {
                  row: { type: "integer" },
                  text: { type: "string" },
                  amount: { anyOf: [{ type: "number" }, { type: "null" }] }
                },
                required: ["row", "text", "amount"],
                additionalProperties: false
              }
            },
            fill_remaining: { type: "boolean" }
          },
          required: ["items", "fill_remaining"],
          additionalProperties: false
        }
      }
    },
    system: "You turn what someone typed into their meal plan into the foods to log from their food list, with how much of each they had. " +
      "Each line of the list is \"row: food name (brand) | serving amount\". The typed text may be abbreviated, misspelled, missing the brand, " +
      "or worded differently from the list. It may name one food, or something made of several foods (e.g. \"half pb sandwich\" is bread " +
      "plus peanut butter); return one item per separate food, in the order they'd be logged. " +
      "For each item, set row to the row number of the list food it most likely is, or -1 if nothing in the list is a plausible match. " +
      "If the typed text names a brand, store or restaurant for a food, only match a list item from that same brand (or one with no brand that is " +
      "clearly the same product); a similar food from a different brand is not a match, so use -1. " +
      "Set text to a short description of that one food on its own, including its brand, its quantity, and any instruction that applies to it such as " +
      "\"conservative estimate\" (it is used to look the food up online when row is -1). " +
      "Set amount to the item's quantity expressed in the unit of the chosen row's serving amount: a bare number means that unit already " +
      "(\"popcorn 8\" with a serving of \"30g\" is 8), and a quantity in a different unit is converted (0.3 oz is 8.5 g). For count-based servings " +
      "like \"1 Cup\" or \"7 Almonds\", give the number of those units, and use the food's typical density if a weight-volume conversion is needed. " +
      "If a single food is typed on its own with no quantity, set amount to null. For each part of a combined item (like the bread and peanut butter " +
      "in a sandwich), always set amount: use the quantity typed for that part, or estimate a typical one (the bread in half a sandwich is one slice). " +
      "Amount must always be in the unit of the row's serving amount, so if the serving is \"1g\" or \"100g\", amount is a number of grams " +
      "(one slice of whole wheat bread is about 40, not 1); only count-based servings like \"1 slice\" take a count. When row is -1, set amount to null. " +
      "If the typed text is gibberish or names no food, return an empty items list. " +
      "If instead the typed text asks you to fill the rest of the day's calories (e.g. \"fill my remaining cals with some foods i like\"), " +
      "set fill_remaining to true and return an empty items list; otherwise set fill_remaining to false.",
    messages: [{ role: "user", content: "Food list:\n" + foodList + "\nTyped text: " + enteredText }]
  });
  if (result === null || !Array.isArray(result.items)) return null;
  Logger.log("LLM chose " + JSON.stringify(result.items) + " for: " + enteredText);
  const items = [];
  for (const item of result.items) {
    if (item.row === -1) {
      items.push({ row: -1, text: String(item.text), amount: null });
    } else if (Number.isInteger(item.row) && item.row >= 1 && item.row < data.length) {
      const amount = typeof item.amount === "number" && item.amount > 0 ? item.amount : null;
      items.push({ row: item.row, text: String(item.text), amount: amount });
    }
  }
  return { fillRemaining: result.fill_remaining === true, items: items };
}

const CONSERVATIVE_MACRO_FACTOR = 0.7; // Asking for a conservative estimate cuts the looked-up macros by 30%

// For a food that isn't in the Foods tab, has Claude search the web for its nutrition facts.
// Returns {food, brand, serving, protein, carbs, fat, amount} (amount is the typed quantity in the
// serving's unit, or null), or null if the text isn't a recognizable food or the call failed.
// If the typed text asks for a conservative estimate, the macros are multiplied by CONSERVATIVE_MACRO_FACTOR.
function lookUpNewFoodWithLlm(enteredText) {
  const result = callClaudeForJson({
    model: LLM_MODEL,
    max_tokens: 16000,
    fallbacks: "default",
    output_config: {
      effort: "high",
      format: {
        type: "json_schema",
        schema: {
          type: "object",
          properties: {
            is_food: { type: "boolean" },
            food: { type: "string" },
            brand: { type: "string" },
            serving: { type: "string" },
            protein: { type: "number" },
            carbs: { type: "number" },
            fat: { type: "number" },
            amount: { anyOf: [{ type: "number" }, { type: "null" }] },
            conservative: { type: "boolean" }
          },
          required: ["is_food", "food", "brand", "serving", "protein", "carbs", "fat", "amount", "conservative"],
          additionalProperties: false
        }
      }
    },
    tools: [{ type: "web_search_20260209", name: "web_search", max_uses: 5 }],
    system: "You estimate nutrition facts for a food someone typed into their meal plan, so it can be added to their food list. " +
      "First decide whether the typed text names a recognizable food or drink (a generic food, a dish, or a branded or restaurant product). " +
      "If it is gibberish or not a food, set is_food to false and leave the other fields empty or zero. " +
      "Otherwise you must use web search before answering, even if you think you know the numbers. Look up reliable nutrition facts " +
      "(the brand's or restaurant's own nutrition info first, then USDA or another reputable database) and give your best estimate. " +
      "Use a clear, specific food name (title case, without the brand) and the brand if there is one, or an empty string. " +
      "Choose a sensible serving: the package or menu serving for branded and restaurant items, or a gram weight such as 100g for generic foods. " +
      "Write serving like the existing list does, e.g. \"28g\", \"1 Bowl\", \"150 mL\", \"7 Almonds\". Give protein, carbs and fat in grams for that serving. " +
      "If the typed text includes a quantity, set amount to that quantity expressed in the serving's unit (a bare number means grams for foods " +
      "measured by weight); otherwise set amount to null. " +
      "Set conservative to true if the typed text asks for a conservative estimate (e.g. \"conservative\", \"be conservative\", \"cons est\"), " +
      "otherwise false. Don't treat that request as part of the food's name, and report the nutrition facts as you found them; " +
      "the adjustment is applied afterwards.",
    messages: [{ role: "user", content: "Typed text: " + enteredText }]
  });
  if (result === null) return null;
  Logger.log("New food lookup for " + enteredText + ": " + JSON.stringify(result));
  if (!result.is_food || String(result.food).trim() === "" || isNaN(parseFloat(result.serving))) return null;
  const amount = typeof result.amount === "number" && result.amount > 0 ? result.amount : null;
  const factor = result.conservative === true ? CONSERVATIVE_MACRO_FACTOR : 1;
  if (factor !== 1) Logger.log("Conservative estimate requested: multiplying macros by " + factor);
  return {
    food: result.food, brand: result.brand, serving: result.serving,
    protein: parseFloat((result.protein * factor).toFixed(2)),
    carbs: parseFloat((result.carbs * factor).toFixed(2)),
    fat: parseFloat((result.fat * factor).toFixed(2)),
    amount: amount
  };
}

// Takes a Foods row (Food, Brand, Amount, Protein, Carbs, Fat, ...) and returns its first six values with the
// amount set to newNum (in the row's existing unit) and the macros scaled to match, the same way
// changeMacrosBasedOnChangeToAmount does when the amount is edited. Returns null if the amount can't be scaled.
function scaleFoodRow(foodRow, newNum) {
  const oldVal = String(foodRow[2]);
  if (oldVal.includes("/")) return null; // Fraction amounts aren't scaled, as in changeMacrosBasedOnChangeToAmount
  const oldNum = parseFloat(oldVal);
  if (isNaN(oldNum) || oldNum === 0) return null;
  const unitChars = oldVal.match(/(?<=\d+)\D*$/) || "";
  const multiplier = newNum / oldNum;
  const row = foodRow.slice(0, 6);
  row[2] = parseFloat(newNum.toFixed(2)) + unitChars;
  for (let col = 3; col <= 5; col++) {
    const macroVal = parseFloat(String(row[col]));
    if (!isNaN(macroVal)) row[col] = parseFloat((macroVal * multiplier).toFixed(2));
  }
  Logger.log("Scaled " + row[0] + " from " + oldVal + " to " + row[2]);
  return row;
}

// Run once from the editor to replace the simple onEdit trigger with an installable one, and to schedule
// the nightly snapshot.
function installEditTrigger() {
  for (const t of ScriptApp.getProjectTriggers()) {
    if (["onEditInstalled", "saveDailySnapshot"].includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  }
  ScriptApp.newTrigger("onEditInstalled").forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  ScriptApp.newTrigger("saveDailySnapshot").timeBased().everyDays(1).atHour(23).nearMinute(55).create();
}

// The first row goes in startRow and the rest in the blank rows below it in the same meal.
function writeRowsIntoMeal(mealPlanSheet, startRow, mealRows) {
  const targetRows = [startRow];
  const lastRow = mealPlanSheet.getLastRow();
  if (mealRows.length > 1 && lastRow > startRow) {
    const below = mealPlanSheet.getRange(startRow + 1, 1, lastRow - startRow, 7).getValues();
    for (let i = 0; i < below.length && targetRows.length < mealRows.length; i++) {
      if (String(below[i][0]).trim() !== "") break; // Column A marks the Meal Total row or the next meal
      if (below[i].slice(1, 7).every(x => String(x).trim() === "")) targetRows.push(startRow + 1 + i);
    }
  }
  for (let i = 0; i < mealRows.length; i++) {
    if (i >= targetRows.length) {
      Logger.log("No blank row left in this meal for: " + mealRows[i][0]);
      continue;
    }
    mealPlanSheet.getRange(targetRows[i], 2, 1, 6).setValues([mealRows[i]]);
  }
}

const HISTORY_SPREADSHEET_PROPERTY = "HISTORY_SPREADSHEET_ID";

// Returns the first sheet of the "Meal Plan History" spreadsheet (its ID is kept in script properties),
// creating the spreadsheet in My Drive if createIfMissing is true. Returns null if there isn't one.
function getHistorySheet(createIfMissing) {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(HISTORY_SPREADSHEET_PROPERTY);
  if (id) {
    try {
      return SpreadsheetApp.openById(id).getSheets()[0];
    } catch (err) {
      Logger.log("Couldn't open the history spreadsheet: " + err);
    }
  }
  if (!createIfMissing) return null;
  const file = SpreadsheetApp.create("Meal Plan History");
  const sheet = file.getSheets()[0];
  sheet.getRange(1, 1, 1, 9).setValues([["Date", "Meal", "Food", "Brand", "Amount", "Protein", "Carbs", "Fat", "Calories"]]);
  sheet.setFrozenRows(1);
  props.setProperty(HISTORY_SPREADSHEET_PROPERTY, file.getId());
  Logger.log("Created Meal Plan History: " + file.getUrl());
  return sheet;
}

// Returns the Meal Plan rows inside a numbered meal section ("1 - Breakfast" through its Meal Total), each as
// {meal, values}. Rows in other sections, like the Remainder list below the meals, are left out.
function mealSectionRows(plan) {
  const rows = [];
  let meal = null;
  for (let i = 1; i < plan.length; i++) {
    const label = String(plan[i][0]).trim();
    if (label === "Meal Total") continue;
    if (/^\d+\s*-/.test(label)) meal = label;
    else if (label !== "") meal = null;
    if (meal !== null) rows.push({ meal: meal, values: plan[i] });
  }
  return rows;
}

function formatHistoryDate(value, timeZone) {
  return value instanceof Date ? Utilities.formatDate(value, timeZone, "yyyy-MM-dd") : String(value);
}

// Runs nightly from a time-based trigger (see installEditTrigger). Appends the day's filled-in Meal Plan rows to the
// Meal Plan History spreadsheet, replacing any rows already saved for the same date.
function saveDailySnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timeZone = TIME_ZONE;
  const date = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
  const plan = ss.getSheetByName("Meal Plan").getDataRange().getValues();
  const rows = [];
  for (const { meal, values: r } of mealSectionRows(plan)) {
    if (String(r[1]).trim() === "" || String(r[3]).trim() === "") continue; // Only rows with a food and an amount
    rows.push([date, meal, r[1], r[2], String(r[3]), r[4], r[5], r[6], r[7]]);
  }
  const sheet = getHistorySheet(true);
  const existing = sheet.getDataRange().getValues();
  for (let i = existing.length - 1; i >= 1; i--) {
    if (formatHistoryDate(existing[i][0], timeZone) === date) sheet.deleteRow(i + 1);
  }
  if (rows.length === 0) return;
  const firstRow = sheet.getLastRow() + 1;
  sheet.getRange(firstRow, 1, rows.length, 1).setNumberFormat("@"); // Keep the date and amount as plain text
  sheet.getRange(firstRow, 5, rows.length, 1).setNumberFormat("@");
  sheet.getRange(firstRow, 1, rows.length, 9).setValues(rows);
  Logger.log("Saved " + rows.length + " rows for " + date);
}

// When each meal starts (hour of the day, 24h) and a word that appears in its Meal Plan section name. Filling the
// remaining calories favors foods usually eaten at the meal for the time the request is entered.
const MEAL_START_HOURS = [[0, "dinner"], [4, "breakfast"], [12, "lunch"], [14, "afternoon"], [20, "dinner"]];
const TIME_ZONE = "America/Denver"; // Mountain Time, for meal times and snapshot dates

function mealForNow() {
  const hour = parseInt(Utilities.formatDate(new Date(), TIME_ZONE, "H"), 10);
  let meal = MEAL_START_HOURS[0][1];
  for (const [startHour, name] of MEAL_START_HOURS) if (hour >= startHour) meal = name;
  return meal;
}

// Fills what's left of the day (the Remainder row under the Target row: protein, carbs, fat and calories) with foods
// usually eaten at the current meal (by time of day), drawn from the Meal Plan History, today's plan, and the
// "I like" lists below the Remainder row. Starts at startRow and continues into the blank rows below it in the same meal.
function fillRemainingCalories(ss, startRow, data) {
  const mealPlanSheet = ss.getSheetByName("Meal Plan");
  const plan = mealPlanSheet.getDataRange().getValues();
  const remainderIndex = plan.findIndex(r => String(r[0]).trim() === "Remainder");
  if (remainderIndex < 0) {
    Logger.log("No Remainder row found on the Meal Plan");
    return;
  }
  const remaining = plan[remainderIndex].slice(4, 8).map(x => parseFloat(x)); // Protein, carbs, fat, calories
  if (isNaN(remaining[3]) || remaining[3] <= 0) {
    Logger.log("No remaining calories to fill: " + remaining[3]);
    return;
  }
  const timeZone = TIME_ZONE;
  const meal = mealForNow();
  const foodIndex = {};
  function foodKey(food, brand) {
    return String(food).trim().toLowerCase() + "|" + String(brand).trim().toLowerCase();
  }
  for (let i = 1; i < data.length; i++) {
    const key = foodKey(data[i][0], data[i][1]);
    if (!(key in foodIndex)) foodIndex[key] = i;
  }
  // For each Foods row: the days it was eaten (at any meal, and at this meal), the amounts, and the "I like" lists it's on
  const stats = {};
  function statsFor(row) {
    if (!stats[row]) stats[row] = { days: {}, mealDays: {}, amounts: [], mealAmounts: [], liked: false, likedForMeal: false };
    return stats[row];
  }
  function countFood(day, mealLabel, food, brand, amount) {
    const row = foodIndex[foodKey(food, brand)];
    if (row === undefined) return;
    const s = statsFor(row);
    const atThisMeal = String(mealLabel).toLowerCase().includes(meal);
    s.days[day] = true;
    if (atThisMeal) s.mealDays[day] = true;
    if (String(amount).trim() === "") return;
    s.amounts.push(String(amount));
    if (atThisMeal) s.mealAmounts.push(String(amount));
  }
  const historySheet = getHistorySheet(false);
  if (historySheet !== null) {
    const history = historySheet.getDataRange().getValues();
    for (let i = 1; i < history.length; i++) {
      countFood(formatHistoryDate(history[i][0], timeZone), history[i][1], history[i][2], history[i][3], history[i][4]);
    }
  }
  for (const { meal: mealLabel, values: r } of mealSectionRows(plan)) countFood("today", mealLabel, r[1], r[2], r[3]);
  let likedListForMeal = false;
  for (let i = remainderIndex + 1; i < plan.length; i++) {
    const text = String(plan[i][1]).trim();
    if (/i like:?$/i.test(text)) { // A list heading like "Breakfast I like:"
      likedListForMeal = text.toLowerCase().includes(meal);
      continue;
    }
    const row = foodIndex[foodKey(plan[i][1], plan[i][2])];
    if (row === undefined) continue;
    const s = statsFor(row);
    s.liked = true;
    if (likedListForMeal) s.likedForMeal = true;
    if (String(plan[i][3]).trim() !== "") (likedListForMeal ? s.mealAmounts : s.amounts).push(String(plan[i][3]));
  }
  const count = obj => Object.keys(obj).length;
  const score = row => {
    const s = stats[row];
    return 3 * count(s.mealDays) + (s.likedForMeal ? 3 : 0) + count(s.days) + (s.liked ? 1 : 0);
  };
  const candidates = Object.keys(stats).map(Number).sort((a, b) => score(b) - score(a)).slice(0, 40);
  if (candidates.length === 0) {
    Logger.log("No frequently eaten or liked foods found to fill with");
    return;
  }
  let candidateList = "";
  for (const i of candidates) {
    const d = data[i];
    const s = stats[i];
    const usual = (s.mealAmounts.length > 0 ? s.mealAmounts : s.amounts).slice(-3).join(", ");
    candidateList += i + ": " + d[0] + (d[1] ? " (" + d[1] + ")" : "") + " | serving " + d[2] + " | P/C/F " + d[3] + "/" + d[4] + "/" + d[5] +
      " | " + d[6] + " kcal per serving | eaten at " + meal + " on " + count(s.mealDays) + " days, at any meal on " + count(s.days) + " days" +
      (s.likedForMeal ? " | on their " + meal + " 'I like' list" : s.liked ? " | on an 'I like' list for another meal" : "") +
      " | usual amounts: " + usual + "\n";
  }
  const result = callClaudeForJson({
    model: LLM_MODEL,
    max_tokens: 16000,
    fallbacks: "default",
    output_config: {
      effort: "high",
      format: {
        type: "json_schema",
        schema: {
          type: "object",
          properties: {
            items: {
              type: "array",
              items: {
                type: "object",
                properties: { row: { type: "integer" }, amount: { type: "number" } },
                required: ["row", "amount"],
                additionalProperties: false
              }
            }
          },
          required: ["items"],
          additionalProperties: false
        }
      }
    },
    system: "You plan the rest of someone's day of eating while they are on a calorie-controlled cut. Choose 2 to 5 foods from the candidates, " +
      "with amounts, so that what you add fits what they have left for the day: land as close as possible to the remaining calories (aim for " +
      "within 25 kcal and never more than 50 over), then get as close as you can to the remaining protein, carbs and fat (a negative number means " +
      "they're already over, so keep that macro low). Strongly favor foods they usually eat at the current meal (eaten at that meal on many days, " +
      "or on their 'I like' list for that meal), and only reach for foods from other meals if those can't fit. Choose combinations they'd " +
      "plausibly eat together at that meal, with realistic portions close to their usual amounts. Each candidate line is " +
      "\"row: food (brand) | serving | P/C/F per serving | kcal per serving | days eaten at this meal and at any meal | 'I like' lists | usual amounts\". " +
      "Give amount in the unit of the serving: grams if the serving is in grams, or a count of units for servings like \"1 Scoop\" or " +
      "\"7 Almonds\". A food's calories and macros are amount / (the serving's number) × the per-serving values. Add up your totals before " +
      "answering and adjust the amounts until they fit.",
    messages: [{
      role: "user",
      content: "Current meal: " + meal + "\nRemaining for today: " + remaining[3] + " kcal, protein " + remaining[0] + " g, carbs " + remaining[1] +
        " g, fat " + remaining[2] + " g\nCandidate foods:\n" + candidateList
    }]
  });
  if (result === null || !Array.isArray(result.items)) return;
  const mealRows = [];
  const added = [0, 0, 0];
  for (const item of result.items) {
    if (!candidates.includes(item.row) || !(item.amount > 0)) continue;
    const foodRow = data[item.row].slice(0, 6);
    const scaledRow = scaleFoodRow(foodRow, item.amount);
    const row = scaledRow !== null ? scaledRow : foodRow;
    for (let m = 0; m < 3; m++) added[m] += parseFloat(row[3 + m]) || 0;
    mealRows.push(row);
  }
  Logger.log("Filling " + meal + " with " + remaining.join("/") + " left (P/C/F/kcal); adding " + added.map(x => x.toFixed(1)).join("/") + "/" +
    Math.round(added[0] * 4 + added[1] * 4 + added[2] * 9) + ": " + JSON.stringify(mealRows));
  writeRowsIntoMeal(mealPlanSheet, startRow, mealRows);
}
