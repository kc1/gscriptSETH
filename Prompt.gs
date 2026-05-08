function promptStage1() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('OPTIMIZER');
  let currentSheetObjArr = sheet2Json(sheet);
  Logger.log("length: " + currentSheetObjArr.length);

  // Filter rows that haven't been processed yet (no RoadAvailable prompt? Wait, adjust filter as needed)
  // Assuming you want rows where GeneratedAnswer is empty
  const toProcess = currentSheetObjArr.filter(row => {
    return row.PROMPT &&
      row.NealNotes &&
      row.RoadURL.includes("http") &&
      (!row.RoadAvailable || String(row.RoadAvailable).trim() === "");
  });

  Logger.log(`Rows to process: ${toProcess.length}`);

  // Process first 10 (or all if you prefer)
  const firstXPushedRows = toProcess.slice(0, 10);

  if (firstXPushedRows.length === 0) {
    Logger.log("No new rows to process.");
    return;
  }

  Logger.log(JSON.stringify(firstXPushedRows));

  let url = APIURL + 'openRouterPromptRun1';
  Logger.log(url);

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(firstXPushedRows),
    muteHttpExceptions: true
  };


  Logger.log('payload');

  Logger.log(JSON.stringify(firstXPushedRows));


  try {
    // Make the API request
    const response = UrlFetchApp.fetch(url, options);

    Logger.log(response)

    // Parse the JSON response if it is JSON
    var result = JSON.parse(response.getContentText());

    // Log the result
    Logger.log(result.message);

    const filteredRows = JSON.parse(result.results);


    for (var i = 0; i < filteredRows.length; i++) {

      const myRow = filteredRows[i];
      Logger.log(myRow.RoadAvailable);

      // const APN = result.message.APN;
      // const APN2 = result.message.APN2;
      // const GEOM = result.message.GEOM;

      // var AN = updateCell(sheet, myRow, 'ContourResponse', myRow.ContourResponse);
      // var AE = updateCell(sheet, myRow, 'WaterResponse', myRow.WaterResponse);

      let Response = myRow.RoadAvailable;

      let rResponse = Response.split('').reverse().join('');

      rResponse = rResponse.substr(0, rResponse.indexOf("---"));

      Response = rResponse.split('').reverse().join('');

      Logger.log(Response);


      let json = "{" + extractSubstring(Response, "{", "}") + "}";


      Logger.log(json);

      const obj = JSON.parse(json);


      let YN = obj["RoadAvailable"];



      var C = updateCell(sheet, myRow, 'RoadAvailable', YN);


    }

  } catch (error) {
    // Log any errors
    Logger.log(error);
  }


}

function stage1Status() {

  const out = statusCalculator("RoadAvailable", "NealNotes", "STATUS");
  Logger.log(out);

}


function statusCalculator(column1, column2, outputColumn) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('OPTIMIZER');
  let currentSheetObjArr = sheet2Json(sheet);
  Logger.log("length: " + currentSheetObjArr.length);

  let matchNum = 0
  for (var i = 0; i < 91; i++) {
    const myRow = currentSheetObjArr[i];
    if (myRow[column1].length == myRow[column2].length) {
      var C = updateCell(sheet, myRow, outputColumn, "MATCH");
      matchNum++;
    } else {
      var C = updateCell(sheet, myRow, outputColumn, "MISMATCH");

    }

  }
  const myRow = currentSheetObjArr[92];
  var C = updateCell(sheet, myRow, outputColumn, matchNum + " MATCHS");


  return "FIN"
}


// ====================== RUN THIS AFTER Status IS UPDATED ======================
function refineMismatchedPrompts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  let data = sheet2Json(sheet);

  // Find mismatches
  const mismatches = data.filter(row =>
    String(row.Status).trim().toUpperCase() === "MISMATCH" &&
    row.RoadURL &&
    row.PROMPT &&
    row.NealNotes
  );

  Logger.log(`Found ${mismatches.length} mismatches to refine.`);

  if (mismatches.length === 0) {
    Logger.log("No mismatches found.");
    return;
  }


  for (var i = 0; i < mismatches.length; i++) {

    const myRow = mismatches[i];
    Logger.log(myRow.RoadAvailable);
    Logger.log(`Refining prompt for Row ID: ${myRow.ID}`);
    const improvedPrompt = callPromptOptimizer(myRow);

    if (improvedPrompt && improvedPrompt.trim() !== "") {
      const newVersion = incrementVersion(myRow.PromptVersion);

      var A = updateCell(sheet, myRow, 'PROMPT', improvedPrompt);
      var B = updateCell(sheet, myRow, 'RoadAvailable', "");
      var C = updateCell(sheet, myRow, 'Status', "");
      var D = updateCell(sheet, myRow, 'PromptVersion', improvedPrompt);

      Logger.log(`✓ Row ${myRow.ID} prompt updated to ${newVersion}`);
    } else {
      Logger.log(`✗ Failed to get improved prompt for Row ${myRow.ID}`);
    }
  }

  Logger.log("Prompt refinement completed.");
}

function incrementVersion(current) {
  if (!current || current === "v0") return "v1";
  const num = parseInt(current.replace("v", "")) || 0;
  return `v${num + 1}`;
}

// Call this from your serverless function
function callPromptOptimizer(row) {
  const payload = {
    task: "refine_prompt",
    imageUrl: row.RoadURL,
    currentPrompt: row.PROMPT,
    wrongAnswer: row.RoadAvailable,
    correctAnswer: row["Correct answer"] || "unknown",   // Add this column if missing
    rowId: row.ID
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  Logger.log('payload');
  Logger.log(payload);

  try {
    const res = UrlFetchApp.fetch(APIURL + 'openRouterRoadAvailable', options);
    const data = JSON.parse(res.getContentText());
    return data.improvedPrompt;
  } catch (e) {
    Logger.log("Optimizer failed: " + e);
    return null;
  }
}


// ====================== RUN THIS AFTER Status IS UPDATED ======================
function refineMismatchedPrompts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  let data = sheet2Json(sheet);

  // Find mismatches
  const mismatches = data.filter(row =>
    row.Status.includes("MISMATCH") &&
    row.RoadURL &&
    row.PROMPT &&
    row.NealNotes
  );

  Logger.log(`Found ${mismatches.length} mismatches to refine.`);

  if (mismatches.length === 0) {
    Logger.log("No mismatches found.");
    return;
  }


  for (var i = 0; i < mismatches.length; i++) {

    const myRow = mismatches[i];
    const improvedPrompt = callPromptOptimizer(myRow);
    if (improvedPrompt && improvedPrompt.trim() !== "") {
      const newVersion = incrementVersion(row.PromptVersion);
      var C = addNote2(sheet, myRow, 'PROMPT', improvedPrompt);
      var C = addNote2(sheet, myRow, 'RoadAvailable', "");
      var C = addNote2(sheet, myRow, 'Status', "");
      var C = addNote2(sheet, myRow, 'PromptVersion', newVersion);

    }
  }

  Logger.log("Prompt refinement completed.");
}


function incrementVersion(currentVersion) {
  if (!currentVersion || currentVersion === "") return "v1";
  const match = currentVersion.match(/v(\d+)/);
  const num = match ? parseInt(match[1]) : 0;
  return `v${num + 1}`;
}

