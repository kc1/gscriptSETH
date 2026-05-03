function onOpen() {
  // var ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  const menu = [

    { name: '1) Push to Named Bucket (up to 50), for home access-then run with IMAGEAI to generate screenshots', functionName: 'alcornPush' },
    , { name: '2) Run sethProp (runSethApp) with IMAGEAI to generate screenshots --api updated', functionName: 'postToSethProp' },

    { name: '3) Update selected rows with Screenshot Links from Named Mongo bucket (alcornBucket) ', functionName: 'updateWithScreenshotPaths' },

    // { name: '4A) GeoJSONio url builder from collection', functionName: 'GeoJSONioUrlBuilder' },

    { name: '4A) Push Road Rows to collection', functionName: 'alcornGeoJsonPush' },

    { name: '4B) Run buildScreenshotsFromLink (runSethApp) with IMAGEAI to generate screenshots ', functionName: 'postToBuildScreenshotsFromLink' },

    { name: '4C) Update rows with Road Screenshot Links from collection (alcornGeoJsonBucket) ', functionName: 'updateWithRoadScreenshotPaths' },

    { name: '4D)  Y/N on Available Road using WaterURL with LLM  -- works in parallel', functionName: 'roadAvailableUsingLLM' },



    { name: '5) Auto Process Screenshot Links with LLM3 -- works in parallel', functionName: 'autoToLLM3' },
    { name: '6) Calculate Points in multiple Rows', functionName: 'getPointsInMultipleRows' },
    { name: '7) Auto Process Screenshot Links for Frontage -- works in parallel', functionName: 'getFrontageWithLLM' },




    { name: '7) Select rows and Run STR Calculation (Please switch to the STR sheet)', functionName: 'STR' },

    { name: '8) Get Single or Multiple APNs and WKTs - api updated', functionName: 'getAPNS' },

    { name: '10) If its still interesting Send Row to Pipeline', functionName: 'toPipeline' },
    { name: 'Distance to Closest Walmart', functionName: 'getWallmartDistance' },

    { name: '-----------------------------------------------------------------------', functionName: 'dummy' },



    { name: 'Load Planning/Zoning data', functionName: 'loadPZ' },
    { name: 'Initial Planning/Zoning review request (email county PZ office)', functionName: 'initialMessagePZ' },
    { name: 'Create KML file in Drive - then upload to GEOJSON.IO for map creation', functionName: 'buildKML' },
    { name: 'Modify Colours and titles of geojson file (copy geojson.io data into A10)', functionName: 'markupGeojson' },

    { name: 'Create a Planning/Zoning map review request message in Message field (for email only) - Need to have kml file ready for review', functionName: 'createMessagePZ' },

    { name: 'Select Row and send PZ department an email', functionName: 'sendEmailToPZFromRow' },
    // { name: '-----------------------------------------------------------------', functionName: 'dummy' },

  ]
  ss.addMenu('Extended', menu);

  // const menu2 = [];

  const menu3 = [
    { name: 'Add Agent to Follow Up', functionName: 'followAgent' },

    { name: 'Update Remote Mongo Atlas DB from Land.com', functionName: 'updateDBNewLandcomWi' },
    { name: 'GET LAND.COM - wisconsin properties', functionName: 'getNewLandcomWi' },
    { name: 'Create Realtor Message', functionName: 'RealMessage' },

    // { name: 'Build Google Earth URL', functionName: 'buildGearthURL' },
    // { name: 'Create Wetland mapper URL', functionName: 'WetlandUrlBuilder' },
    // { name: 'FEMA FLOODplane mapper URL', functionName: 'addFloodPlainLink' },
    { name: 'Build BB from Polygon', functionName: 'buildBB' },
    // { name: 'Build url to download DEM geotiff from TNM', functionName: 'createTnmUrl' },
    // { name: 'Import Datatree Property Json', functionName: 'importJsonToSheet' },
    // { name: 'Import Land.com Property Json', functionName: 'importJsonFromDoc' },
    // { name: 'Backup GEO', functionName: 'getGEO' },
    { name: 'Get Agent Data', functionName: 'getAgent' },
    { name: 'Set todays OUTREACH contacts to PENDING (9-10 daily)', functionName: 'setTodaysContactsToPending' },
    { name: 'Get todays PENDING and email (10-11 daily)', functionName: 'SendTodaysOutreachEmails' },
    { name: 'Extract Utilities from Detail Page', functionName: 'extractUtilities' },

  ];
  // ss.addMenu('RESEARCH', menu2);
  ss.addMenu('Extras', menu3);


}

const APIURL = "https://nimble-dieffenbachia-92e8e2.netlify.app/.netlify/functions/";


function tester() {
  startup();
  Utilities.sleep(210 * 1000)
  getNewWIProps();
}
function dummy() {
}

function postToSethProp() {
  var url = 'https://image1.space/sethProp';

  var payload = {
    num: 30,
    filterObj: {}
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getResponseCode());
  Logger.log(response.getContentText());
}

function postToBuildScreenshotsFromLink() {
  var url = 'https://image1.space/buildScreenshotsFromLink';

  var payload = {
    num: 30,
    filterObj: {}
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getResponseCode());
  Logger.log(response.getContentText());
}

function markupGeojson() {
  // Sample GeoJSON data (use your own variable)
  var PIPELINESheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PIPELINE');
  PIPELINESheet.getRange("D10").setValue("");

  var value = PIPELINESheet.getRange("A10").getValue();
  Logger.log(value);


  const geojson = JSON.parse(value);

  // Helper function to convert HSL to RGB string
  function hslToRgbString(h, s, l) {
    // Convert h, s, l (0-1 range)
    s /= 100;
    l /= 100;
    let c = (1 - Math.abs(2 * l - 1)) * s,
      x = c * (1 - Math.abs((h / 60) % 2 - 1)),
      m = l - c / 2;
    let r = 0, g = 0, b = 0;
    if (h < 60) { r = c; g = x; b = 0; }
    else if (h < 120) { r = x; g = c; b = 0; }
    else if (h < 180) { r = 0; g = c; b = x; }
    else if (h < 240) { r = 0; g = x; b = c; }
    else if (h < 300) { r = x; g = 0; b = c; }
    else { r = c; g = 0; b = x; }
    r = Math.round((r + m) * 255);
    g = Math.round((g + m) * 255);
    b = Math.round((b + m) * 255);
    return `RGB (${r}, ${g}, ${b})`;
  }

  // Get RGB color with distinct separation for each lot
  function getColor(idx, total) {
    // Distribute hues evenly, use 65% saturation, 60% lightness for nice colors
    let hue = (idx * 360 / total) % 360;
    return hslToRgbString(hue, 65, 60);
  }

  // Assign fill property with RGB string
  geojson.features.forEach((feature, idx) => {
    feature.properties.name = `Lot ${idx}`;

    feature.properties.fill = getColor(idx, geojson.features.length);

  });


  PIPELINESheet.getRange("D10").setValue(JSON.stringify(geojson));



}


function returnRecentResponses() {
  // Get received emails from the last 7 days
  const now = new Date();
  const threeDaysAgo = new Date(now.getTime() - (7 * 24 * 60 * 60 * 1000));

  // Convert timestamps (milliseconds → seconds)
  const newerThan = Math.floor(threeDaysAgo.getTime() / 1000);
  const query = "newer:" + newerThan + " in:inbox category:primary";

  // Search for threads
  const threads = GmailApp.search(query);
  Logger.log("Found " + threads.length + " email threads in the last 3 days.");

  if (threads.length === 0) {
    Logger.log("No recent email responses found.");
    return;
  }

  // Process each thread
  const emailResponses = [];

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = thread.getMessages();

    // Get the latest message in the thread (most recent response)
    const latestMessage = messages[messages.length - 1];

    Logger.log(" latest thread message");
    Logger.log(latestMessage);

    // Skip if this is a sent message (not a received response)
    if (latestMessage.getFrom().includes(Session.getActiveUser().getEmail())) {
      continue;
    }

    const subject = latestMessage.getSubject();
    const from = latestMessage.getFrom();
    const body = latestMessage.getPlainBody();
    const date = latestMessage.getDate();

    if (!from.includes("google.com")) {


      // Create a single file content with all thread information
      let fileContent = `Subject: ${subject}\n`;
      fileContent += `From: ${from}\n`;
      fileContent += `Date: ${date}\n`;
      fileContent += `Thread ID: ${thread.getId()}\n`;
      fileContent += `---\n\n`;
      fileContent += `Body:\n${body}\n\n`;

      // Add all messages in the thread for context
      fileContent += `--- THREAD CONTEXT ---\n\n`;
      for (let j = 0; j < messages.length; j++) {
        const msg = messages[j];
        fileContent += `Message ${j + 1}:\n`;
        fileContent += `From: ${msg.getFrom()}\n`;
        fileContent += `Date: ${msg.getDate()}\n`;
        fileContent += `Subject: ${msg.getSubject()}\n`;
        fileContent += `Body: ${msg.getPlainBody()}\n\n`;
      }

      // Create response object
      const responseObj = {
        subject: subject,
        from: from,
        body: body,
        date: date,
        threadId: thread.getId(),
        fileContent: fileContent,
        processedAt: new Date()
      };

      Logger.log('response obj ' + i);
      Logger.log(responseObj);

      emailResponses.push(responseObj);
    }
  }

  Logger.log(emailResponses);

  return emailResponses


}

function getSignature() {


  return "<div>" +
    "<div>" + "Dan Nir" + "</div>" +
    "<div>" + "The Open Land company" + "</div>" +
    "<div>" + "508-296-5085" + "</div>" +
    "</div>";

  // MailApp.sendEmail({
  //   to: "recipient@example.com",
  //   subject: "Subject goes here",
  //   body: "This is the plain text body.",
  //   htmlBody: "<p>This is the message body.</p>" + signature
  // });
}


function loadPZ() {

  const myRow = getSelectedRowObject();
  var COUNTYSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('COUNTY INFO');


  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  const countySheetObjArr = sheet2Json(COUNTYSheet);
  // Logger.log("length: " + currentSheetObjArr.length);
  const county = myRow.county;
  const countyRowArr = countySheetObjArr.filter(row => row.COUNTY == county)
  const countyRow = countyRowArr.pop();
  Logger.log(countyRow);
  const pzObj = JSON.parse(countyRow.ZONINGDEPT);
  // PlanZoneName	PlanZoneEmail
  updateCell(sheet, myRow, 'PlanZoneName', pzObj.title + " " + pzObj.ln);
  updateCell(sheet, myRow, 'PlanZoneEmail', pzObj.email);


}






function calculatePointsInRow(roadJson, contourJson, waterJson) {

  Logger.log("RR");
  Logger.log(roadJson);
  // CRobj = JSON.parse(CR);
  // Logger.log(CRobj)

  roadJson = roadJson.replace(/'/g, '"');
  contourJson = contourJson.replace(/'/g, '"');
  waterJson = waterJson.replace(/'/g, '"');


  const RR = JSON.parse(roadJson); // {'roadNumberInteger': 2}
  const CR = JSON.parse(contourJson); //{estimated percentage of lot that is hilly: 10, estimated percentage of lot that is flat: 90}
  const WR = JSON.parse(waterJson); // {estimated percentage flood zone: 0, estimated percentage ground water: 0, total estimated percentage: 0}

  let roadNumberInteger;
  let flatLand;
  let waterValue;

  switch (true) {
    case RR["roadNumberInteger"] === 0:
      roadNumberInteger = 0;
      break;
    case RR["roadNumberInteger"] === 1:
      roadNumberInteger = 50;
      break;
    case RR["roadNumberInteger"] > 1:
      roadNumberInteger = 65;
      break;
    default:
      // optionally handle other cases
      break;
  }
  flatLand = Number(CR["estimated percentage of lot that is flat"]) / 100 * 20;
  waterValue = (100 - Number(WR["total estimated percentage"])) / 100 * 50;

  // Logger.log("CR");
  // Logger.log(CR);
  // CRobj = JSON.parse(CR);
  // Logger.log(CRobj);


  Logger.log("roadNumberInteger");
  Logger.log(roadNumberInteger);

  Logger.log("flatLand");
  Logger.log(flatLand);

  Logger.log("waterValue");
  Logger.log(waterValue);

  return roadNumberInteger + flatLand + waterValue;
}


function getScreenshotPaths() {

  // const payload = getMultipleSelectedRowObjects();
  const payload = getMultipleSelectedRowObjectsDiscontinuous();

  // const payload = getSelectedRowObject();

  Logger.log(payload);
  Logger.log(typeof (payload))

  Logger.log(JSON.stringify(payload))
  Logger.log("-------------------------");

  const url = "https://comfy-crisp-d74946.netlify.app/.netlify/functions/getDropboxLinks";

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    // Make the API request
    const response = UrlFetchApp.fetch(url, options);

    // Parse the JSON response if it is JSON
    const responseObjArr = JSON.parse(response.getContentText());

    // Log the result
    Logger.log('out');
    Logger.log(responseObjArr);
    const bucket_listing_ids = responseObjArr.map(x => x.listing_id);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
    const currentSheetObjArr = sheet2Json(sheet);
    Logger.log("length: " + currentSheetObjArr.length);


    // for (var i = 0; i < currentSheetObjArr.length; i++) { //right now its > 2000 - I'll scan first 300 properties
    for (var i = 0; i < 300; i++) { //right now its > 2000 - I'll scan first 300 properties
      const myRow = currentSheetObjArr[i];

      // const currentSheetObjArr = objArr[i];
      if (bucket_listing_ids.includes(myRow.listing_id)) {

        let filtered = responseObjArr.filter(row => row.listing_id === myRow.listing_id);
        Logger.log(filtered);

        const filteredRow = filtered.pop();

        var waterFile = filteredRow.WaterURL
        var contourFile = filteredRow.ContourURL

        Logger.log("Water File: " + waterFile);
        Logger.log("Contour File: " + contourFile);

        // var C = updateCell(sheet, myRow, 'ContourURL', contourFile);
        var C = addNote2(sheet, myRow, 'ContourURL', contourFile);

        var W = addNote2(sheet, myRow, 'WaterURL', waterFile);
        // var C = addNote2(sheet, myRow, column, text);

      } else {

        var C = addNote2(sheet, myRow, 'ContourURL', "REPUSH!");

        var W = addNote2(sheet, myRow, 'WaterURL', "Return to STEP 4");
      }
    }

  } catch (error) {
    // Log any errors
    Logger.log(error);
  }


}

function updateWithScreenshotPaths() {

  autoUpdateWithScreenshotPaths("alcornBucket", 20);

}

function updateWithRoadScreenshotPaths() {

  const isPushed = (x) => x.RoadURL === "PUSHED";
  baseUpdateWithScreenshotPaths("alcornGeoJsonBucket", 20, isPushed);

}

function autoUpdateWithScreenshotPaths(collectionName, numberToPull) {

  // const responseObjArr = getMultipleSelectedRowObjects();
  // const responseObjArr = getMultipleSelectedRowObjectsDiscontinuous();
  // const payload = getSelectedRowObject();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  let currentSheetObjArr = sheet2Json(sheet);
  Logger.log("length: " + currentSheetObjArr.length);


  const pushedRows = currentSheetObjArr.filter(x => {
    if (x.ContourURL == "PUSHED" || x.WaterURL == "PUSHED") { return x }

  })

  let firstFiftyPushedRows = pushedRows.slice(0, numberToPull);// get first 50 or less

  Logger.log('out');
  const IDS = firstFiftyPushedRows.map(x => x.ID);
  Logger.log(IDS);


  const filterObj = { ID: { $in: IDS } };

  const myJSON = fetchMongoDBDataAPI(filterObj, collectionName)

  Logger.log(myJSON);
  const myObjects = JSON.parse(myJSON);
  let mongoRecords = myObjects.documents;

  let rowUploadedToBucketFlag = true; // STEP 4
  let rowUpdatedWithLinksFlag = true; // STEP 5


  for (var i = 0; i < firstFiftyPushedRows.length; i++) { //right now its > 2000 - I'll scan first 300 properties
    // for (var i = 0; i < 300; i++) { //right now its > 2000 - I'll scan first 300 properties
    const myRow = firstFiftyPushedRows[i];

    // const currentSheetObjArr = objArr[i];
    // if (listing_ids.includes(myRow.listing_id)) {

    let filteredRecords = mongoRecords.filter(mongoRow => mongoRow.ID === myRow.ID);
    Logger.log(filteredRecords);

    if (filteredRecords.length) {

      const filteredRecord = filteredRecords.pop(); // get the first , likely only record.

      var waterFile = filteredRecord?.WaterURL
      var contourFile = filteredRecord?.ContourURL

      if (waterFile.startsWith("http") || contourFile.startsWith("http")) { // ie UPDATED step 5.

        // Logger.log("Water File: " + waterFile);
        // Logger.log("Contour File: " + contourFile);

        // var C = updateCell(sheet, myRow, 'ContourURL', contourFile);
        if (contourFile) {
          var C = updateCell(sheet, myRow, 'ContourURL', contourFile);
        }
        if (waterFile) {
          var W = updateCell(sheet, myRow, 'WaterURL', waterFile);
        } else {
          var C = updateCell(sheet, myRow, 'ContourURL', "");
          var W = updateCell(sheet, myRow, 'WaterURL', "");
          rowUpdatedWithLinksFlag = false;
        }
      }
    } else {
      var C = updateCell(sheet, myRow, 'ContourURL', "");
      var W = updateCell(sheet, myRow, 'WaterURL', "");
      rowUploadedToBucketFlag = false;
    }


  }
  if (!rowUploadedToBucketFlag) {
    SpreadsheetApp.getUi().alert("You may need to rerun step 4!");
  }
  if (!rowUpdatedWithLinksFlag) {
    SpreadsheetApp.getUi().alert("You may need to rerun step 5!");
  }
}



function baseUpdateWithScreenshotPaths(collectionName, numberToPull, filterFunction) {

  // const responseObjArr = getMultipleSelectedRowObjects();
  // const responseObjArr = getMultipleSelectedRowObjectsDiscontinuous();
  // const payload = getSelectedRowObject();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  let currentSheetObjArr = sheet2Json(sheet);
  Logger.log("length: " + currentSheetObjArr.length);


  const pushedRows = currentSheetObjArr.filter(filterFunction);

  let firstXPushedRows = pushedRows.slice(0, numberToPull);// get first 50 or less

  Logger.log('out');
  const IDS = firstXPushedRows.map(x => x.ID);
  Logger.log(IDS);


  const filterObj = { ID: { $in: IDS } };

  const myJSON = fetchMongoDBDataAPI(filterObj, collectionName)

  Logger.log(myJSON);
  const myObjects = JSON.parse(myJSON);
  let mongoRecords = myObjects.documents;

  let rowUploadedToBucketFlag = true; // STEP 4
  let rowUpdatedWithLinksFlag = true; // STEP 5


  for (var i = 0; i < firstXPushedRows.length; i++) { //right now its > 2000 - I'll scan first 300 properties
    // for (var i = 0; i < 300; i++) { //right now its > 2000 - I'll scan first 300 properties
    const myRow = firstXPushedRows[i];

    // const currentSheetObjArr = objArr[i];
    // if (listing_ids.includes(myRow.listing_id)) {

    let filteredRecords = mongoRecords.filter(mongoRow => mongoRow.ID === myRow.ID);
    Logger.log(filteredRecords);

    if (filteredRecords.length) {

      const filteredRecord = filteredRecords.pop(); // get the first , likely only record.

      var roadFile = filteredRecord?.RoadURL

      if (roadFile.startsWith("http")) { // ie UPDATED step 5.
        var C = updateCell(sheet, myRow, 'RoadURL', roadFile);
      }
    } else {
      var C = updateCell(sheet, myRow, 'RoadURL', "");
    }

  }

}



// function addNote2(text = "", column = "NOTES") {

function getNote2(row, column) {

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // const myRow = getSelectedRowObject();
  const cr = getRowCellRangeByColumnName(row, column);
  console.log(cr);

  // var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get the sheet by name
  var sheet = spreadsheet.getSheetByName('Sheet1');

  // Get the specific cell (e.g., A1)
  var cell = sheet.getRange(cr);

  // Retrieve the existing note
  var currentNote = cell.getNote();
  return currentNote;
}


function addNote2(sheet, row, column, text) {


  // const myRow = getSelectedRowObject();
  const cr = getRowCellRangeByColumnName(row, column);
  console.log(cr);

  // var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get the sheet by name
  // var sheet = spreadsheet.getSheetByName('Sheet1');

  // Get the specific cell (e.g., A1)
  var cell = sheet.getRange(cr);

  // Retrieve the existing note
  // var currentNote = cell.getNote();

  // Define the new bullet point to add
  // var newBulletPoint = "• " + text;

  // Append the new bullet point to the existing note
  // var updatedNote = currentNote + "\n" + newBulletPoint;

  // Update the cell with the new note
  cell.setNote(text);

}

function prependNote2(sheet, row, column, text) {


  // const myRow = getSelectedRowObject();
  const cr = getRowCellRangeByColumnName(row, column);
  console.log(cr);

  // var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get the sheet by name
  // var sheet = spreadsheet.getSheetByName('Sheet1');

  // Get the specific cell (e.g., A1)
  var cell = sheet.getRange(cr);

  // Retrieve the existing note
  var currentNote = cell.getNote();

  // Define the new bullet point to add
  var newBulletPoint = text + "======================";

  // Append the new bullet point to the existing note
  var updatedNote = newBulletPoint + "\n" + currentNote;

  // Update the cell with the new note
  cell.setNote(updatedNote);

}



function autoToLLM3() {
  // Get selected rows
  // const rowObjArr = getMultipleSelectedRowObjectsDiscontinuous();
  // const listing_ids = rowObjArr.map(x => x.listing_id);

  // const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  // const currentSheetObjArr = sheet2Json(sheet);


  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  let currentSheetObjArr = sheet2Json(sheet);
  Logger.log("length: " + currentSheetObjArr.length);

  // let firstTen = currentSheetObjArr.slice(0, 10);// get first 10 or less



  const pushedRows = currentSheetObjArr.filter(x => {
    if ((x.ContourURL.includes("dropbox") && x.WaterURL.includes("dropbox")) && (x.ContourResponse == "" || x.WaterResponse == "" || x.RoadResponse == "" || x.ContourResponse == "{}" || x.WaterResponse == "{}" || x.RoadResponse == "{}")) { return x }
  })

  const filteredRows = pushedRows.filter(x => { return x.ID && x.ContourURL && x.WaterURL && x.ContourResponse && x.WaterResponse && x.RoadResponse && x.POINTS });

  let firstTenPushedRows = pushedRows.slice(0, 15);// get first 10 or less

  Logger.log(JSON.stringify(firstTenPushedRows));

  const keepKeys = ['ID', 'ContourURL', 'WaterURL', 'ContourResponse', 'WaterResponse', 'RoadResponse', 'POINTS'];

  const limitedRows = firstTenPushedRows.map(row =>
    Object.fromEntries(
      Object.entries(row).filter(([key]) => keepKeys.includes(key))
    )
  );

  Logger.log(JSON.stringify(limitedRows));


  let url = APIURL + 'openRouter';

  Logger.log(JSON.stringify(firstTenPushedRows));

  // let url ="https://www.postb.in/1775252069951-5364527481142";
  // url = "https://www.postb.in/b/1775252069951-5364527481142";
  Logger.log(url);

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(firstTenPushedRows),
    muteHttpExceptions: true
  };


  Logger.log('payload');

  Logger.log(JSON.stringify(options))


  try {
    // Make the API request
    var response = UrlFetchApp.fetch(url, options);
    Logger.log("Response Code: " + response.getResponseCode());
    Logger.log("Response: " + response.getContentText());
    // Parse the JSON response if it is JSON
    var result = JSON.parse(response.getContentText());

    // Log the result
    Logger.log(result.message);

    const filteredRows = JSON.parse(result.results);


    for (var i = 0; i < filteredRows.length; i++) {

      const myRow = filteredRows[i];

      // const APN = result.message.APN;
      // const APN2 = result.message.APN2;
      // const GEOM = result.message.GEOM;

      // var AN = updateCell(sheet, myRow, 'ContourResponse', myRow.ContourResponse);
      // var AE = updateCell(sheet, myRow, 'WaterResponse', myRow.WaterResponse);

      let ContourResponse = myRow.ContourResponse;
      let WaterResponse = myRow.WaterResponse;
      let RoadResponse = myRow.RoadResponse;

      let rContourResponse = ContourResponse.split('').reverse().join('');
      let rWaterResponse = WaterResponse.split('').reverse().join('');
      let rRoadResponse = RoadResponse.split('').reverse().join('');

      rContourResponse = rContourResponse.substr(0, rContourResponse.indexOf("---"));
      rWaterResponse = rWaterResponse.substr(0, rWaterResponse.indexOf("---"));
      rRoadResponse = rRoadResponse.substr(0, rRoadResponse.indexOf("---"));

      ContourResponse = rContourResponse.split('').reverse().join('');
      WaterResponse = rWaterResponse.split('').reverse().join('');
      RoadResponse = rRoadResponse.split('').reverse().join('');

      Logger.log(ContourResponse);
      Logger.log(WaterResponse);
      Logger.log(RoadResponse);


      let contourJson = "{" + extractSubstring(ContourResponse, "{", "}") + "}";
      let waterJson = "{" + extractSubstring(WaterResponse, "{", "}") + "}";
      let roadJson = "{" + extractSubstring(RoadResponse, "{", "}") + "}";


      Logger.log(contourJson);
      Logger.log(waterJson);
      Logger.log(roadJson);

      // const points = calculatePointsInRow(roadJson, contourJson, waterJson)


      if (contourJson.length > 10) {
        var C = updateCell(sheet, myRow, 'ContourResponse', contourJson);
      }
      if (waterJson.length > 10) {
        var W = updateCell(sheet, myRow, 'WaterResponse', waterJson);
      }
      if (roadJson.length > 10) {
        var Y = updateCell(sheet, myRow, 'RoadResponse', roadJson);
      }
      // var Z = updateCell(sheet, myRow, 'POINTS', points);


    }

  } catch (error) {
    // Log any errors
    Logger.log("Error: " + error.toString());
  }


}

function getPointsInMultipleRows() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  let currentSheetObjArr = sheet2Json(sheet);
  Logger.log("length: " + currentSheetObjArr.length);


  const pushedRows = currentSheetObjArr.filter(x =>
    typeof x.ContourResponse === 'string' &&
    typeof x.WaterResponse === 'string' &&
    typeof x.RoadResponse === 'string' &&
    x.ContourResponse.length > 2 &&
    x.WaterResponse.length > 2 &&
    x.RoadResponse.length > 2 &&
    typeof x.POINTS !== 'number'
  );

  let firstTenPushedRows = pushedRows.slice(0, 50);// get first 10 or less
  Logger.log(JSON.stringify(firstTenPushedRows));



  for (var i = 0; i < firstTenPushedRows.length; i++) {

    const myRow = firstTenPushedRows[i];
    const contourJson = myRow.ContourResponse;
    const waterJson = myRow.WaterResponse;
    const roadJson = myRow.RoadResponse;

    Logger.log(contourJson);
    Logger.log(waterJson);
    Logger.log(roadJson);
    const points = calculatePointsInRow(roadJson, contourJson, waterJson)
    var Z = updateCell(sheet, myRow, 'POINTS', points);

  }


}

function roadAvailableUsingLLM() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  let currentSheetObjArr = sheet2Json(sheet);
  Logger.log("length: " + currentSheetObjArr.length);

  const pushedRows = currentSheetObjArr.filter(x => {
    if (x.WaterURL.includes("dropbox") && !String(x.RoadAvailable || "")) { return x }
  })

  // x.WaterURL.includes("box") && !String(x.available || "").trim()



  let firstTenPushedRows = pushedRows.slice(0, 10);// get first 10 or less

  Logger.log(JSON.stringify(firstTenPushedRows));




  let url = APIURL + 'openRouterRoadAvailable';
  Logger.log(url);

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(firstTenPushedRows),
    muteHttpExceptions: true
  };


  Logger.log('payload');

  Logger.log(JSON.stringify(options));


  try {
    // Make the API request
    const response = UrlFetchApp.fetch(url, options);

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


      let YN = obj["AvailableRoad"];



      var C = updateCell(sheet, myRow, 'RoadAvailable', YN);


    }

  } catch (error) {
    // Log any errors
    Logger.log(error);
  }


}

function getFrontageWithLLM() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  let currentSheetObjArr = sheet2Json(sheet);
  Logger.log("length: " + currentSheetObjArr.length);

  const pushedRows = currentSheetObjArr.filter(x => {
    if (x.WaterURL.includes("dropbox") && x.calculatedPerimeterFeet !== "" && x.calcFrontage == "") { return x }
    // if ( x.calcFrontage == "") { return x }

  })

  let firstTenPushedRows = pushedRows.slice(0, 10);// get first 10 or less

  Logger.log(JSON.stringify(firstTenPushedRows));

  const keepKeys = ['ID', 'ContourURL', 'WaterURL', 'ContourResponse', 'WaterResponse', 'RoadResponse', 'POINTS'];


  // const limitedRows = firstTenPushedRows.map(row =>
  //   Object.fromEntries(
  //     Object.entries(row).filter(([key]) => keepKeys.includes(key))
  //   )
  // );

  // Logger.log(JSON.stringify(limitedRows));


  let url = APIURL + 'openRouterFrontage';
  // let url ="https://www.postb.in/1775252069951-5364527481142";
  // url = "https://www.postb.in/b/1775252069951-5364527481142";
  Logger.log(url);

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(firstTenPushedRows),
    muteHttpExceptions: true
  };


  Logger.log('payload');

  Logger.log(JSON.stringify(options));


  try {
    // Make the API request
    const response = UrlFetchApp.fetch(url, options);

    // Parse the JSON response if it is JSON
    var result = JSON.parse(response.getContentText());

    // Log the result
    Logger.log(result.message);

    const filteredRows = JSON.parse(result.results);


    for (var i = 0; i < filteredRows.length; i++) {

      const myRow = filteredRows[i];
      Logger.log(myRow.calcFrontage);

      // const APN = result.message.APN;
      // const APN2 = result.message.APN2;
      // const GEOM = result.message.GEOM;

      // var AN = updateCell(sheet, myRow, 'ContourResponse', myRow.ContourResponse);
      // var AE = updateCell(sheet, myRow, 'WaterResponse', myRow.WaterResponse);

      let FrontageResponse = myRow.calcFrontage;

      let rFrontageResponse = FrontageResponse.split('').reverse().join('');

      rFrontageResponse = rFrontageResponse.substr(0, rFrontageResponse.indexOf("---"));

      FrontageResponse = rFrontageResponse.split('').reverse().join('');

      Logger.log(FrontageResponse);


      let frontageJson = "{" + extractSubstring(FrontageResponse, "{", "}") + "}";


      Logger.log(frontageJson);

      const frontageObj = JSON.parse(frontageJson);


      frontage = frontageObj["Estimated Road Length in feet"];



      var C = updateCell(sheet, myRow, 'calcFrontage', frontage);


    }

  } catch (error) {
    // Log any errors
    Logger.log(error);
  }


}


function callOpenRouter2(imageLink, prompt) {
  var url = "https://openrouter.ai/api/v1/chat/completions";

  var headers = {
    "Authorization": "Bearer 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",
    "Content-Type": "application/json"
  };

  // "model": "meta-llama/llama-3.2-90b-vision-instruct",


  var payload = {
    "model": "google/gemini-2.5-flash",
    "messages": [
      {
        "role": "user",
        "content": [
          {
            "type": "text",
            "text": prompt // No need to manually add quotes
          },
          {
            "type": "image_url",
            "image_url": {
              "url": imageLink // Use the passed parameter
            }
          }
        ]
      }
    ]
  };

  var options = {
    method: "post",
    headers: headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // Allows handling of non-200 responses
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var result = response.getContentText();
    Logger.log("Response: " + result);
    return result;
  } catch (error) {
    Logger.log("Error: " + error);
  }
}


function imageTest() {

  const text = "What is in this image?";
  // const imageLink = 'https://www.dropbox.com/scl/fi/b7tfviennf40hgz5c5lrn/wi_Marathon_7830082140992_contours.png?rlkey=plpddg83p82dxcebndlg9xlla&dl=0';
  const imageLink = 'https://www.dropbox.com/scl/fi/b7tfviennf40hgz5c5lrn/wi_Marathon_7830082140992_contours.png?rlkey=plpddg83p82dxcebndlg9xlla&raw=1';


  let x = callOpenRouter2(imageLink, text);


}

function alcornPush() {

  resp = pushToNamedBucket('alcornBucket', 30);

}

function alcornGeoJsonPush() {

  resp = geoJsonPush('alcornGeoJsonBucket', 20);

}


function pushToNamedBucket(collection, numberToPush) {

  // - NOTE TO SEND TO netlify update Bucket function it must be an array with a property_id 

  // const payload = getMultipleSelectedRowObjectsDiscontinuous();
  // const payload = getMultipleSelectedRowObjects();
  // const payload = getSelectedRowObject();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  // const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet2');



  const currentSheetObjArr = sheet2Json(sheet);
  Logger.log("length: " + currentSheetObjArr.length);


  const payload = currentSheetObjArr.filter(x => {
    if (x.ContourURL == "" || x.WaterURL == "") { return x }

  })
  let firstFiftyPushedRows = payload.slice(0, numberToPush);// get first 50 or less


  // Updated payload: Wrap the array and collection name in an object
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      data: firstFiftyPushedRows,  // The array of objects
      collectionName: collection  // The collection name as a string
    }),
    muteHttpExceptions: true
  };


  Logger.log(firstFiftyPushedRows);
  Logger.log(typeof (firstFiftyPushedRows))

  // const url = 'https://www.postb.in/1739670263592-0602835712488';
  let url = 'https://nimble-dieffenbachia-92e8e2.netlify.app/.netlify/functions/updateNamedBucket';

  Logger.log(url);


  Logger.log(options);


  try {
    // Make the API request
    const response = UrlFetchApp.fetch(url, options);

    // Parse the JSON response if it is JSON
    var result = JSON.parse(response.getContentText());

    // Log the result

    const myRows = result.message;
    Logger.log(myRows);
    for (var i = 0; i < myRows.length; i++) {

      const myRow = myRows[i];
      var AN = updateCell(sheet, myRow, 'ContourURL', "PUSHED");
      var AE = updateCell(sheet, myRow, 'WaterURL', "PUSHED");


    }


    // const APN = result.message.APN;
    // const APN2 = result.message.APN2;
    // const GEOM = result.message.GEOM;

    // var AN = updateCell(sheet, myRow, 'APN', APN);
    // var AE = updateCell(sheet, myRow, 'APN2', APN2);
    // var AP = updateCell(sheet, myRow, 'GEOM', GEOM);

  } catch (error) {
    // Log any errors
    Logger.log(error);
  }

}

function geoJsonPush(collection, numberToPush) {

  // - NOTE TO SEND TO netlify update Bucket function it must be an array with a property_id 

  // const payload = getMultipleSelectedRowObjectsDiscontinuous();
  // const payload = getMultipleSelectedRowObjects();
  // const payload = getSelectedRowObject();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  // const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet2');



  const currentSheetObjArr = sheet2Json(sheet);
  Logger.log("length: " + currentSheetObjArr.length);


  const rows = currentSheetObjArr.filter(x => {
    if (x.RoadURL == "") { return x }

  })

  // const miniRows = rows.map(x => ({
  //   PARNO: x.PARNO,
  //   RoadURL: x.RoadURL
  // }));

  let firstXRows = rows.slice(0, numberToPush);// get first 50 or less


  // Updated payload: Wrap the array and collection name in an object
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      data: firstXRows,  // The array of objects
      collectionName: collection  // The collection name as a string
    }),
    muteHttpExceptions: true
  };


  Logger.log(firstXRows);
  Logger.log(typeof (firstXRows))

  // const url = 'https://www.postb.in/1739670263592-0602835712488';
  let url = 'https://nimble-dieffenbachia-92e8e2.netlify.app/.netlify/functions/updateNamedBucket';

  Logger.log(url);


  Logger.log(options);


  try {
    // Make the API request
    const response = UrlFetchApp.fetch(url, options);

    // Parse the JSON response if it is JSON
    var result = JSON.parse(response.getContentText());

    // Log the result

    const myRows = result.message;
    Logger.log(myRows);
    for (var i = 0; i < myRows.length; i++) {

      const myRow = myRows[i];
      var AN = updateCell(sheet, myRow, 'RoadURL', "PUSHED");
      // var AE = updateCell(sheet, myRow, 'WaterURL', "PUSHED");


    }


    // const APN = result.message.APN;
    // const APN2 = result.message.APN2;
    // const GEOM = result.message.GEOM;

    // var AN = updateCell(sheet, myRow, 'APN', APN);
    // var AE = updateCell(sheet, myRow, 'APN2', APN2);
    // var AP = updateCell(sheet, myRow, 'GEOM', GEOM);

  } catch (error) {
    // Log any errors
    Logger.log(error);
  }

}

function autoPushToBucket() {

  // - NOTE TO SEND TO netlify update Bucket function it must be an array with a property_id 

  // const payload = getMultipleSelectedRowObjectsDiscontinuous();
  // const payload = getMultipleSelectedRowObjects();
  // const payload = getSelectedRowObject();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  const currentSheetObjArr = sheet2Json(sheet);
  Logger.log("length: " + currentSheetObjArr.length);


  const payload = currentSheetObjArr.filter(x => {
    if (x.AutoEvaluation == "Y" && (x.ContourURL == "" || x.WaterURL == "")) { return x }

  })
  let firstFiftyPushedRows = payload.slice(0, 50);// get first 50 or less

  Logger.log(firstFiftyPushedRows);
  Logger.log(typeof (firstFiftyPushedRows))

  // const url = 'https://www.postb.in/1739670263592-0602835712488';
  let url = 'https://nimble-dieffenbachia-92e8e2.netlify.app/.netlify/functions/updateBucket';

  Logger.log(url);

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(firstFiftyPushedRows),
    muteHttpExceptions: true
  };

  Logger.log(options);


  try {
    // Make the API request
    const response = UrlFetchApp.fetch(url, options);

    // Parse the JSON response if it is JSON
    var result = JSON.parse(response.getContentText());

    // Log the result

    const myRows = result.message;
    Logger.log(myRows);
    for (var i = 0; i < myRows.length; i++) {

      const myRow = myRows[i];
      var AN = updateCell(sheet, myRow, 'ContourURL', "PUSHED");
      var AE = updateCell(sheet, myRow, 'WaterURL', "PUSHED");


    }


    // const APN = result.message.APN;
    // const APN2 = result.message.APN2;
    // const GEOM = result.message.GEOM;

    // var AN = updateCell(sheet, myRow, 'APN', APN);
    // var AE = updateCell(sheet, myRow, 'APN2', APN2);
    // var AP = updateCell(sheet, myRow, 'GEOM', GEOM);

  } catch (error) {
    // Log any errors
    Logger.log(error);
  }

}



function toPipeline() {

  const myObj = getSelectedRowObject();
  var pipelineSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PIPELINE');
  const o = addRowFromObject(myObj, pipelineSheet);

}

function askOpenRouterAI(email) {
  // Define your API key and other configurations
  const OPENROUTER_API_KEY = "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";
  // const YOUR_SITE_URL = "your-site-url-here";     // Replace with your site URL
  // const YOUR_SITE_NAME = "your-site-name-here";   // Replace with your site name

  // const question = "Can you rewrite the following email: Hi Michael, We talked a while back about https://www.realtor.com/realestateandhomes-detail/Campbell-Rd_Herbster_WI_54844_M92817-42723?from=srp , but someone else ended up buying it. Have you come across any similar properties recently in Bayfield? I'm interested in properties with a mild slope , a small amount of ground water/flood plain and legal road access. I can make an all cash offer and move quickly. Thank you - Dan"


  const question = "With the role of a real estate investor, can you rewrite the following email,including the provided url and return the subject and body as json using html syntax: " + email;

  // API endpoint
  const url = "https://openrouter.ai/api/v1/chat/completions";

  // Request payload
  const payload = {
    model: "google/learnlm-1.5-pro-experimental:free",
    messages: [
      { role: "user", content: question } // Dynamically insert user's question
    ],
    top_p: 1,
    temperature: 1,
    repetition_penalty: 1
  };

  // Request headers
  const headers = {
    Authorization: `Bearer ${OPENROUTER_API_KEY}`,
    "Content-Type": "application/json"
  };

  // Request options
  const options = {
    method: "post",
    headers: headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    // Make the HTTP POST request
    const response = UrlFetchApp.fetch(url, options);

    // Parse the JSON response
    const jsonResponse = JSON.parse(response.getContentText());

    // Log the full response for debugging
    Logger.log("Full Response: " + JSON.stringify(jsonResponse));

    // Extract and return the AI's response
    const aiResponse = jsonResponse.choices[0]?.message?.content || "No response received.";
    return aiResponse;

  } catch (error) {
    // Log and handle errors
    Logger.log("Error: " + error.message);
    return "An error occurred while communicating with the AI.";
  }
}


function SendTodaysOutreachEmails() {

  // eg 35, 70, 105 days later ...
  // relies only on follow_up_date in Main

  Logger.log(' IN longTermFollowUp');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outreachSheet = ss.getSheetByName('Outreach');
  var rows = sheet2Json(outreachSheet);

  var filteredRows = rows.filter(function (row) {
    return row.NextContact == "PENDING";
  });

  Logger.log(filteredRows)

  var today = new Date();
  var todayString = Utilities.formatDate(today, 'America/New_York', "yyyy-MM-dd");



  for (var i = 0; i < filteredRows.length; i++) {

    const myRow = filteredRows[i];
    const randomBetweenMinus5And5 = Math.random() * 10 - 5;
    const daysInFuture = 35 + randomBetweenMinus5And5 // 5 weeks +- 5 days
    const fd = getFutureDate(daysInFuture);
    const ContactedOn = myRow.ContactedOn + ',' + todayString;
    const Response = askOpenRouterAI(myRow.MESSAGE);

    // Extract the JSON content
    const jsonString = Response.replace(/^```json/, '').replace(/```/, '');

    Logger.log(jsonString);

    // Parse the JSON
    const jsonObject = JSON.parse(jsonString);

    console.log(jsonObject.subject); // Output: Following up on Bayfield Property Search
    console.log(jsonObject.body);    // Output: The email body

    try {
      MailApp.sendEmail({
        to: myRow.AgentEmail,
        subject: 'Looking for buildable land',
        htmlBody: jsonObject.body
      });

    }
    catch (err) {
      res = false;
    }
    var ud = updateCell(outreachSheet, myRow, 'NextContact', fd);
    var ud = updateCell(outreachSheet, myRow, 'ContactedOn', ContactedOn);

  }


  // var sendableRows = rows.filter(function (row) {
  //   var rowFollowDate = new Date(row['NEXT_CONTACT'])
  //   var followUpDateString = Utilities.formatDate(rowFollowDate, 'America/New_York', "yyyy-MM-dd");
  //   return followUpDateString == todayString;

  // });

  // if (!sendableRows.length) { Logger.log(' nothing to see here'); return; }

  // var sendableRow = sendableRows[0];
  // Logger.log('sendableRow ' + sendableRow['Phone Number']);

  // var nextDate = getFutureDate(35);
  // var message = 'Hi, I contacted you a while back about a property for sale . Is it still available?';

  // var res = contactByPhoneOrEmail(sendableRow['Phone Number'], sendableRow['Email'], message);

  // if (res) {
  //   var u = updateCell(sheet, sendableRow, 'NEXT_CONTACT', nextDate);
  // } else {
  //   var u = updateCell(sheet, sendableRow, 'NEXT_CONTACT', 'FAILED');
  // }

}

function followAgent() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet1 = ss.getSheetByName('Sheet1');
  const outreachSheet = ss.getSheetByName('Outreach');
  let myObj = getSelectedRowObject();
  myObj.FIRST = myObj.AgentName.split(' ')[0];
  myObj.LAST = myObj.AgentName.split(' ')[1];

  myObj.NextContact = myObj.AgentName.split(' ')[1];


  myObj.MESSAGE = `Hi ${myObj.FIRST}, We talked a while back about ${myObj.AAlink} , but someone else ended up buying it. Have you come across any similar properties recently in ${myObj.county} county? I'm interested in properties with a mild slope , a small amount of ground water/flood plain and legal road access. I can make an all cash offer and move quickly. Thank you - Dan`;

  const randomBetweenMinus5And5 = Math.random() * 10 - 5;

  const daysInFuture = 35 + randomBetweenMinus5And5 // 5 weeks +- 5 days

  myObj.NextContact = getFutureDate(daysInFuture);



  const o = addRowFromObject(myObj, outreachSheet);

  //  COUNTY	AgentName	FIRST	LAST	LINK	AgentPhone	AgentEmail	MESSAGE	SPUNMESSAGE	ContactedOn	DayInterval	NextContact																			



}

function setTodaysContactsToPending() {

  var main_rows = jsonRowsFromSheet('OUTREACH')
  Logger.log(main_rows)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outreachSheet = ss.getSheetByName('Outreach');


  let today = new Date();
  today.setHours(0, 0, 0, 0); // Set time to midnight to compare only the date part

  var filteredRows = main_rows.filter(function (row) {
    var date_time = new Date(row['NextContact']); // Convert to Date object
    date_time.setHours(0, 0, 0, 0); // Set time to midnight

    return date_time.getTime() === today.getTime(); // Compare timestamps
  });

  Logger.log(filteredRows)

  for (var i = 0; i < filteredRows.length; i++) {
    const myRow = filteredRows[i];
    var ud = updateCell(outreachSheet, myRow, 'NextContact', "PENDING");
  }

}


function extractZipCode(url) {
  var regex = /WI_(\d{5})_/;
  var match = url.match(regex);
  if (match) {
    return match[1]; // This will return the 5-digit number
  } else {
    return null; // No match found
  }
}





function createFemaUrl(lat, lon) {
  var baseUrl = "https://msc.fema.gov/portal/search?AddressQuery=";
  var coords = encodeURIComponent(lon + ", " + lat);
  return baseUrl + coords;
}

function addFloodPlainLink() {
  var selection = SpreadsheetApp.getSelection();
  var activerow = selection.getCurrentCell().getRow();
  var sheet = SpreadsheetApp.getActiveSheet();

  const json = sheet2Json(sheet);
  const myRow = json[activerow - 2];

  Logger.log(myRow.lon + ", " + myRow.lat);
  const longUrl = createFemaUrl(myRow.lat, myRow.lon);
  // const respJson = UrlFetchApp.fetch(longUrl);
  // const finalUrl = constructMapServerURL(respJson)
  // Logger.log(finalUrl);

  const clickable = shortenUrl(longUrl);


  var ud = updateCell(sheet, myRow, 'FLOODPLAIN', clickable);
}


function importJSONFromData(jsonObj) {

  // Replace 'Sheet1' with the name of the sheet you want to populate.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');

  // Clear existing data (optional).
  sheet.clear();

  // Define the headers (if needed).
  var headers = Object.keys(data[0]);
  sheet.appendRow(headers);

  // Populate the data.
  for (var i = 0; i < data.length; i++) {
    var row = [];
    for (var j = 0; j < headers.length; j++) {
      row.push(data[i][headers[j]]);
    }
    sheet.appendRow(row);
  }
}



function cosineDistanceBetweenPoints(lat1, lon1, lat2, lon2) {
  const R = 6371e3;
  const p1 = lat1 * Math.PI / 180;
  const p2 = lat2 * Math.PI / 180;
  const deltaP = p2 - p1;
  const deltaLon = lon2 - lon1;
  const deltaLambda = (deltaLon * Math.PI) / 180;
  const a = Math.sin(deltaP / 2) * Math.sin(deltaP / 2) +
    Math.cos(p1) * Math.cos(p2) *
    Math.sin(deltaLambda / 2) * Math.sin(deltaLambda / 2);
  const d = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)) * R;
  return d;
}


function fetchMongoDBDataAPI(filterObj, coll) {

  //  1GxuC9AAuc77xJklnS1PSHVKhZUt3QkcOaqdSYRqoeQVrSnf8jtGtLO3zGlNfm4T -- this is the api auth token and cannot be retrieved

  // NOTE - to get the connection string for remote mongo goto realio .env

  // also wisconsinFS.js is in the prefetchcomps project

  var url = APIURL + "fetchMongoDBDataAPI";

  var headers = {
    "Content-Type": "application/json",
    "Access-Control-Request-Headers": "*"
  };

  var payload = {
    filterObj: filterObj,
    coll: coll
  }

  Logger.log("payload: ");
  Logger.log(JSON.stringify(payload));

  var options = {
    method: "post",
    headers: headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: false
  };

  // const query = { list_date: { $gte: dt }, state: st };


  var response = UrlFetchApp.fetch(url, options);

  if (response.getResponseCode() === 200) {
    var result = response.getContentText();
    Logger.log(result);
    const myObjs = JSON.parse(result);
    Logger.log(myObjs);

    Logger.log(myObjs.documents.length);

    return result;
  } else {
    Logger.log("Error: " + response.getResponseCode());
    Logger.log(response.getContentText());
    return false;
  }
}

function treatData(data) {

  // let treatedArr = [];
  // for (let index = 0; index < objArr.length; index++) {
  // const obj = objArr[index];
  // Preprocess data
  const treated = data.map((obj) => {
    const { coordinate, _id, ...rest } = obj;
    return {
      lat: coordinate ? coordinate.lat : null,
      lon: coordinate ? coordinate.lon : null,
      ...rest,
    };
  });
  console.log("treated ", treated);
  // treatedArr.push(treated);
  // }

  return treated;

}

function proj4() {


  const libraryUrl = "https://cdnjs.cloudflare.com/ajax/libs/proj4js/2.6.3/proj4.min.js";
  const proj4 = eval(UrlFetchApp.fetch(libraryUrl).getContentText());
  proj4();

}

function WetlandUrlBuilder() {

  var selection = SpreadsheetApp.getSelection();
  var activerow = selection.getCurrentCell().getRow();
  var sheet = SpreadsheetApp.getActiveSheet();

  const json = sheet2Json(sheet);
  const myRow = json[activerow - 2];
  // const listing_id = myRow.listing_id;
  // const wkt = myRow.GEOM;
  Logger.log(myRow.lon + ", " + myRow.lat);
  const longUrl = constructGeocodeURL(myRow.lat, myRow.lon);
  const respJson = UrlFetchApp.fetch(longUrl);
  const finalUrl = constructMapServerURL(respJson)
  Logger.log(finalUrl);

  const clickable = shortenUrl(finalUrl);


  var ud = updateCell(sheet, myRow, 'WETLAND', clickable);
  var ud1 = updateCell(sheet, myRow, 'WETLANDcoords', myRow.lon + ", " + myRow.lat);

}

function buildGearthURL() {

  var selection = SpreadsheetApp.getSelection();
  var activerow = selection.getCurrentCell().getRow();
  var sheet = SpreadsheetApp.getActiveSheet();

  const json = sheet2Json(sheet);
  const myRow = json[activerow - 2];
  const listing_id = myRow.listing_id;
  const wkt = myRow.GEOM;

  Logger.log(listing_id);
  const fileName = listing_id + ".kml"

  const kmlString = loadFileContents(fileName)

  const base64encoded = encodeKml(kmlString);
  const encoded = encodeURIComponent(base64encoded);

  Logger.log(encoded);

  const polygon = parseWKTToPolygon(wkt);
  const bb = getBoundingBoxFromPolygonCoordinates(polygon);
  const center = getCenterPointFromBoundingBox(bb);
  console.log(center);

  const base = (lat, lon) => `https://earth.google.com/web/@${lat},${lon},183.18378053a,1000.03477904d,30y,0h,0t,0r/data=`
  const baseString = base(center[0], center[1]);
  const longUrl = `${baseString}${encoded}`;
  console.log(longUrl);

  const clickable = shortenUrl(longUrl);


  var ud = updateCell(sheet, myRow, 'gEARTHurl', longUrl);


}

// please note its  acres: { $gte: minAcreage, $lte: maxAcreage } in land.com


function geoQuery(lon, lat, radiusInMiles, minAcreage, maxAcreage, daysBack = 183) {
  const metersinmile = 1609.34;

  // Calculate the cutoff date
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - daysBack);

  return {
    location: {
      $nearSphere: {
        $geometry: { type: "Point", coordinates: [lon, lat] },
        $maxDistance: radiusInMiles * metersinmile // Convert radius from miles to meters
      }
    },
    // Filter out documents with invalid location fields
    "location.coordinates": { $type: "array", $size: 2 }, // Ensure coordinates array has exactly 2 elements
    "location.coordinates.0": { $type: "number", $ne: null }, // Ensure longitude is a valid number
    "location.coordinates.1": { $type: "number", $ne: null }, // Ensure latitude is a valid number
    lot_acres: { $gte: minAcreage, $lte: maxAcreage },
    sold_date: { $gte: cutoffDate.toISOString() } // Ensure sold_date is within 'daysBack' days from today
  };
}


function getWallmartDistance() {

  const myRow = getSelectedRowObject();

  if (!myRow.lon || !myRow.lat) { throw new Error("Missing lat/long in index property."); return; }

  // Logger.log(myRow);
  // Logger.log(" comps : ");
  // const myQuery = geoQuery(-87.214035,33.802722, 200);
  // const myQuery = geoQuery(Number(myRow.lon), Number(myRow.lat), Number(myRow.RADIUS), Number(myRow.MIN_ACRES), Number(myRow.MAX_ACRES), Number(myRow.DAYS_BACK));
  // const myQuery = { "location": { "$nearSphere": { "$geometry": { "type": "Point", "coordinates": [-87.214035, 33.802722] }, "$maxDistance": 321868 } }, "acres": { "$gte": 5, "$lte": 10 } }

  var myQuery = {
    state: "WI",
    location: {
      $near: {
        $geometry: {
          type: "Point",
          coordinates: [Number(myRow.lon), Number(myRow.lat)]
        },
        $maxDistance: 100000 // optional max distance in meters
      }
    }
  }


  // function cosineDistanceBetweenPoints(lat1, lon1, lat2, lon2) {


  console.log(JSON.stringify(myQuery));


  const myJSON = fetchMongoDBDataAPI(myQuery, "Walmart");
  Logger.log(myJSON);
  const myObjects = JSON.parse(myJSON);
  let walmarts = myObjects.documents;
  for (var i = 0; i < walmarts.length; i++) {
    const walmart = walmarts[i];
    const distance = cosineDistanceBetweenPoints(Number(myRow.lat), Number(myRow.lon), Number(walmart.latitude), Number(walmart.longitude));
    walmarts[i].distance = distance;
  }
  Logger.log(walmarts);

  const closestMetres = walmarts[0].distance;
  const closestMiles = closestMetres / 1000 * 0.621371;
  var sheet = SpreadsheetApp.getActiveSheet();
  var ud = updateCell(sheet, myRow, 'to Walmart(<20)', closestMiles);


}






function getLastFrozenRowHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var lastFrozenRow = sheet.getFrozenRows();
  // var lastFrozenRow = freezeRange[-1];
  var lastFrozenRowHeaders = sheet.getRange(lastFrozenRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  return lastFrozenRowHeaders;
}

function getLastFrozenRowHeadersBySheet(sheet) {

  var lastFrozenRow = sheet.getFrozenRows();
  // var lastFrozenRow = freezeRange[-1];
  var lastFrozenRowHeaders = sheet.getRange(lastFrozenRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  return lastFrozenRowHeaders;
}





function createMessage(type) {

  var sheet = SpreadsheetApp.getActiveSheet();

  const myRow = getSelectedRowObject();
  var myText;
  var column;
  if (type == "R") {

    myText = `${myRow.AgentPhone} Hi, I'm interested in more information about the property in ${myRow.AAlink}. Can you send me the county property number/APN? It looks like there is a road, but is there legal access, power and water? Thank you - Dan`;
    column = "Message";
  } else if (type == "PZ") {
    if (!myRow.MYMAPSHARE || !myRow.PlanZoneName) { throw new Error("Missing PlanZoneName or MYMAPSHARE."); return; }

    myText = `<p>${myRow.PlanZoneName},</p> 
    
    <p>I'm interested in potentially buying and subdividing the property in ${myRow.AAlink}. The address that I have is ${myRow.address}. The APN/Property number is ${myRow.APN}.</p>
    
    <p>Based on my reading of the ${myRow.county} county subdivision docs , I believe the following would be an acceptable minor subdivision : ${myRow.MYMAPSHARE} ? Would you mind taking a look at this map and letting me know what you think , as the expert.</p> <p>Thank you,</p>`;
    column = "PlanZoneMessage";
  } else if (type == "IPZ") {
    if (!myRow.MYMAPSHARE || !myRow.PlanZoneName) { throw new Error("Missing PlanZoneName or MYMAPSHARE."); return; }

    myText = `<p>Hi ${myRow.PlanZoneName},</p>
    
     <p>I'm interested in potentially buying and subdividing some property in your county. I'd like to send you a google map of a preliminary split that I would like to make, just to see if it makes sense to you, before contacting a surveyor. Would you be ok with this? </p>
     
     <p>Thank you,</p>`;
    column = "PlanZoneMessage";

  } else if (type == "IR") {



  }



  // var AN = updateCell(sheet, myRow, 'Message', text);


  // addNote(text = "contacted agent on " + Utilities.formatDate(new Date(), "GMT-4", "dd/MM/yyyy"));
  addNote(text = myText, column);


}


// --- Example Usage ---

function processMyMessage() {
  let fullMessage = "• 6083393388 Hi, I'm interested in more information about the property in https://www.realtor.com/realestateandhomes-detail/19th-Dr_Friendship_WI_53934_M98019-30130?from=srp. Can you send me the county property number/APN? It looks like there is a road, but is there legal access, power and water? Thank you - Dan";
  let markerToFind = "Hi";

  let trimmedMessage = removePrefixUntil(fullMessage, markerToFind);

  Logger.log("Original Message:");
  Logger.log(fullMessage);

  Logger.log("\nTrimmed Message (starting from Hi):");
  Logger.log(trimmedMessage);
}


/**
 * Sends an email to the agent based on the selected row's data.
 * Adds a unique identifier to the email subject to track replies.
 * Updates a note in the sheet to indicate the email was sent.
 */
function sendEmailFromRow(type) {

  const myRow = getSelectedRowObject(); // This function should be defined elsewhere to get data by header name
  // if (!myRow) {
  //   SpreadsheetApp.getUi().alert("Error", "Please select a row with data.", SpreadsheetApp.getUi().ButtonSet.OK);
  Logger.log(myRow);
  // return 
  //   return;
  // }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selectedRowIndex = sheet.getActiveRange().getRow(); // Get the row index of the selected cell
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; // Get headers


  let address;
  let subject;
  let myMessage;

  if (type == "R") {


    myMessage = getNote2(myRow, "Message"); // Assuming getNote2 retrieves note from "Message" column cell

    if (myMessage !== "") {
      myMessage = removePrefixUntil(myMessage, "Hi"); // This function needs to be defined
    }




    address = myRow.AgentEmail;
    subject = myRow.AAlink; // Original subject

    // Generate a unique ID for this email. We'll use the current timestamp plus row index.
    // This helps link replies back to the specific row.
    // Store this ID in a hidden column or as a note if you prefer.
    const uniqueEmailId = "MSG-" + selectedRowIndex + "-" + Date.now();
    subject = subject + " [" + uniqueEmailId + "]"; // Append unique ID to subject

    Logger.log("Sending email to: " + address);
    Logger.log("Subject: " + subject);
    Logger.log("Body:\n" + myMessage);

  } else if (type == "PZ") {
    myMessage = getNote2(myRow, "PlanZoneMessage");

    address = myRow.PlanZoneEmail;
    subject = "Is this property dividable?"
  } else if (type == "IR") {

    subject = "Hi"

    // const myRow = getSelectedRowObject();
    const agentJSON = myRow.INITIALREALTORCANVAS;
    const agentObj = JSON.parse(agentJSON);
    address = agentObj.email;

    myMessage = `<p> Hi ${agentObj.fn},</p>

    <p>I am a land investor and I came across one of your listings, which I thought was pretty well done on land.com. I see that you do a fair amount of land sales . I might be looking for someone to represent me, and I was wondering what counties you work in and know well. </p>

    <p>Thank you, </p>`;

  }


  try {
    MailApp.sendEmail({
      to: address,
      subject: subject,
      // body: myMessage,
      htmlBody: myMessage + getSignature()

      // You can also add Bcc yourself to confirm it was sent, though not strictly necessary for this logic.
      // bcc: Session.getActiveUser().getEmail()
    });
    SpreadsheetApp.getUi().alert("Success", "Email sent to " + address, SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (e) {
    Logger.log("Error sending email: " + e.message);
    SpreadsheetApp.getUi().alert("Error", "Failed to send email: " + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }


}


// function sendEmailFromRow() {

//   const myRow = getSelectedRowObject();
//   let address;
//   let subject;
//   let emailText;

//   let myMessage = getNote2(myRow, "Message");


//   if (myMessage !== "") {

//     emailText = removePrefixUntil(myMessage, "Hi");


//   } else {

//     emailText = "Hi, I'm interested in more information about the property in the title above. Can you send me the county property number/APN? \nIt looks like there is a road, but is there legal access, power and water?\n\nThank you, \n\nDan Nir\n508-296-5085";


//   }
//   address = myRow.AgentEmail;
//   subject = myRow.AAlink;


//   Logger.log(address);
//   Logger.log(subject);
//   Logger.log(emailText);


//   MailApp.sendEmail({
//     to: address,
//     subject: subject,
//     body: emailText,
//   });

//   addNote(text = "contacted agent on " + Utilities.formatDate(new Date(), "GMT-4", "dd/MM/yyyy"));


// }

function getAgent() {


  var selection = SpreadsheetApp.getSelection();
  var activerow = selection.getCurrentCell().getRow();
  var sheet = SpreadsheetApp.getActiveSheet();

  const json = sheet2Json(sheet);
  const myRow = json[activerow - 2];
  // const myRow = getSelectedRowObject();
  const myLink = myRow.AAlink;
  // const encodedLink = encodeURI(myLink);
  // Logger.log(encodedLink);

  // const url = 'https://realio1-c51a04e6b1da.herokuapp.com/getAgent';
  // const url = "https://realio.onrender.com/getAgent";


  // var headers = {
  //   'sec-ch-ua': '"Chromium";v="118", "Google Chrome";v="118", "Not=A?Brand";v="99"',
  //   'sec-ch-ua-mobile': '?0',
  //   'sec-ch-ua-platform': '"Chrome OS"',
  //   'Referer': 'https://realio.onrender.com/',
  //   'Referrer-Policy': 'strict-origin-when-cross-origin',
  //   'Content-Type': 'application/x-www-form-urlencoded', // Add this line for POST requests
  // };

  const url = "https://comfy-crisp-d74946.netlify.app/.netlify/functions/getAgent";
  Logger.log(url);

  var payload = {
    'val': myLink,
  };

  var options = {
    'method': 'post', // Change the method to "post"
    'contentType': 'application/json',
    // 'headers': headers,
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  var response = UrlFetchApp.fetch(url, options);
  var content = response.getContentText();

  // Handle the response content as needed.
  Logger.log(content);

  const myObj = JSON.parse(content);
  Logger.log(myObj);

  const agentDataObj = myObj.message[0];
  const nameStr = agentDataObj.name;
  const emailStr = agentDataObj.emails.join(',');
  const phoneStr = agentDataObj.phones.join(",");

  // myRow.AgentName = nameStr;
  // myRow.AgentEmail = emailStr;
  // myRow.AgentPhone = phoneStr;

  var AN = updateCell(sheet, myRow, 'AgentName', nameStr);
  var AE = updateCell(sheet, myRow, 'AgentEmail', emailStr);
  var AP = updateCell(sheet, myRow, 'AgentPhone', phoneStr);


}

function getAPNS() {

  // const rows = getMultipleSelectedRowObjects();
  const rows = getMultipleSelectedRowObjectsDiscontinuous();

  Logger.log(rows);
  // return;


  // var selection = SpreadsheetApp.getSelection();
  // var activerow = selection.getCurrentCell().getRow();
  var sheet = SpreadsheetApp.getActiveSheet();

  const json = sheet2Json(sheet);

  for (var i = 0; i < rows.length; i += 1) {

    let myRow = rows[i];
    // const myRow = getSelectedRowObject();
    const lat = myRow.lat;
    const lon = myRow.lon;
    //  const listing_id = myRow.listing_id;

    // Logger.log(encodedLink);

    // API endpoint and request options
    // var url = `https://realio1-c51a04e6b1da.herokuapp.com/getAPN?lat=33.129297104&lon=-86.479067`;

    // var url = `https://realio.onrender.com/getAPN?lat=${lat}&lon=${lon}`;

    const url = 'https://comfy-crisp-d74946.netlify.app/.netlify/functions/getAPNS';

    // const url = "https://www.postb.in/1742522889535-0620236727409"


    const payload = {
      "lat": lat,
      "lon": lon,
      "message": "Hello World"
    };

    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload)
    };

    // const responseData = JSON.parse(response.getContentText());
    // Logger.log(responseData);

    // return responseData;

    try {
      // Make the API request
      const response = UrlFetchApp.fetch(url, options);

      // Parse the JSON response if it is JSON
      var result = JSON.parse(response.getContentText());

      // Log the result
      Logger.log(result.message);

      const APN = result.message.APN;
      // const APN2 = result.message.APN2;
      const GEOM = result.message.GEOM;

      var AN = updateCell(sheet, myRow, 'APN', APN);
      // var AE = updateCell(sheet, myRow, 'APN2', APN2);
      var AP = updateCell(sheet, myRow, 'GEOM', GEOM);

    } catch (error) {
      // Log any errors
      Logger.log(error);
    }
  }

}

function getGEO() {

  var selection = SpreadsheetApp.getSelection();
  var activerow = selection.getCurrentCell().getRow();
  var sheet = SpreadsheetApp.getActiveSheet();

  const json = sheet2Json(sheet);
  const myRow = json[activerow - 2];
  // const myRow = getSelectedRowObject();
  const lat = myRow.lat;
  const lon = myRow.lon;

  const url = `https://property-lines.p.rapidapi.com/get_single_us_boundary?coords=${encodeURIComponent(`${lat},${lon}`)}&lat=${encodeURIComponent(`${lat},`)}&lon=${encodeURIComponent(`${lon}`)}`;

  console.log(url);

  // return

  // var url = 'https://property-lines.p.rapidapi.com/get_single_us_boundary?coords=40.79982062892406%2C%20-73.94531505009626&lat=40.79982062892406%2C&lon=-73.94531505009626';

  var options = {
    method: 'GET',
    headers: {
      'X-RapidAPI-Key': '6659718d96msh6e337e09016d37dp1296f0jsnbd289fc2acc9',
      'X-RapidAPI-Host': 'property-lines.p.rapidapi.com'
    }
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var result = response.getContentText();
    Logger.log(result);
  } catch (error) {
    Logger.log(error);
  }

}

function convertGeoJSONToWKT(geoJSON) {
  const coordinates = geoJSON.features[0].geometry.coordinates[0][0];

  // Convert GeoJSON coordinates to a simple polygon coordinate array
  const polygon = coordinates.map(coord => [coord[1], coord[0]]);

  // Convert the polygon array to WKT manually
  const wktString = `POLYGON((${polygon.map(point => point.join(' ')).join(',')}))`;

  return wktString;
}

function buildKML() {

  const sheet = SpreadsheetApp.getActiveSheet();
  const myRow = getSelectedRowObject();
  const wkt = myRow.GEOM;
  const mypolygon = parseWKTToPolygon(wkt)
  Logger.log(mypolygon);
  const title = myRow.address + "," + myRow.county + " county";
  const description = myRow.AAlink;
  createKMLFile(mypolygon, myRow.listing_id, title, description)


}



function wktToGeojson(wkt) {
  var coords = wkt.match(/POLYGON\(\((.*?)\)\)/)[1].split(",");
  var rings = [];
  for (var i = 0; i < coords.length; i += 1) {
    let coordinate = coords[i];
    Logger.log([parseFloat(coords[i]), parseFloat(coords[i + 1])]);
    let coordArr = coordinate.split(" ");
    const coordStr = coordArr.join(",");

    rings.push([parseFloat(coordArr[0]), parseFloat(coordArr[1])]);
  }

  return {
    "type": "Polygon",
    "coordinates": [rings]
  };
}

function normalizeGeoJSON1(input) {
  if (!input || typeof input !== 'object') {
    throw new Error('Input must be an object');
  }

  var geojson = {
    type: input.type || 'Polygon',
    coordinates: input.coordinates || []
  };

  if (input.bbox) {
    geojson.bbox = input.bbox;
  }

  return geojson;
}

function normalizeGeoJSON(obj) {
  if (!obj || typeof obj !== 'object') {
    throw new Error('Input must be an object');
  }

  // Create clean GeoJSON structure
  var geojson = {
    type: obj.type || 'Polygon',
    coordinates: obj.coordinates || []
  };

  // Validate and add bbox if present
  if (Array.isArray(obj.bbox) && obj.bbox.length === 4) {
    geojson.bbox = obj.bbox;
  } else {
    geojson.bbox = calculateBbox(obj.coordinates);
  }

  return geojson;
}

function calculateBbox(coordinates) {
  if (!Array.isArray(coordinates) || coordinates.length === 0) {
    return null;
  }

  var minLng = Infinity, minLat = Infinity, maxLng = -Infinity, maxLat = -Infinity;

  for (var i = 0; i < coordinates.length; i++) {
    var ring = coordinates[i];
    for (var j = 0; j < ring.length; j++) {
      var lng = ring[j][0];
      var lat = ring[j][1];
      minLng = Math.min(minLng, lng);
      minLat = Math.min(minLat, lat);
      maxLng = Math.max(maxLng, lng);
      maxLat = Math.max(maxLat, lat);
    }
  }
  return [minLng, minLat, maxLng, maxLat];
}

function normalizeGeoJSONForGeojsonIO(obj) {
  if (!obj || typeof obj !== 'object') {
    throw new Error('Input must be an object');
  }

  // Create clean geometry
  var geometry = {
    type: obj.type || 'Polygon',
    coordinates: obj.coordinates || []
  };

  // Add bbox if present
  if (Array.isArray(obj.bbox) && obj.bbox.length === 4) {
    geometry.bbox = obj.bbox;
  }

  // Wrap in Feature
  var feature = {
    type: 'Feature',
    geometry: geometry,
    properties: {}  // Empty properties for simple display
  };

  return feature;
}

function GeoJSONioUrlBuilder() {

  // {AAlink=https://www.realtor.com/realestateandhomes-detail/Alma-Rd_Jasper_AL_35501_M98724-43725?from=srp, AgentEmail=, address=Alma Rd, list_date=2024-03-17T23:25:39.000000Z, ACRES/PIECE=, =, flags=is_new_listing, RADIUS=20.0, AgentName=, state=Walker, county=2024-03-18T00:18:19.665Z, APN=, ppa=4000.0, RANGE_PPA=, lot_acres=41.0, updatedAt=al, listing_id=2.965226047E9, price=164000.0, NOTES=, GEOM=POLYGON((-111.179167643 32.5349351874,-111.179166834 32.5358422453,-111.180237626 32.5358413441,-111.180238435 32.5349342937,-111.179167643 32.5349351874)), PIECES=, COMP_AVG_DISTANCE=, relativeRow=4.0, APN2=, MIN_ACRES=5.0, lon=-87.214035, AgentPhone=, MAX_ACRES=10.0, absoluteRow=5.0, lat=33.802722}

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  let currentSheetObjArr = sheet2Json(sheet);
  Logger.log("length: " + currentSheetObjArr.length);

  const pushedRows = currentSheetObjArr.filter(x => {
    if (!String(x.RoadURL || "")) { return x }
  })

  let firstTenPushedRows = pushedRows.slice(0, 3);// get first 10 or less
  Logger.log(JSON.stringify(firstTenPushedRows));


  for (var i = 0; i < firstTenPushedRows.length; i++) {

    let myRow = firstTenPushedRows[i];

    const myParno = myRow.PARNO;

    const filterObj = { PARNO: myParno };

    try {

      const myJSON = fetchMongoDBDataAPI(filterObj, "alcornMERGED2subset");
      Logger.log(myJSON);
      const myObjects = JSON.parse(myJSON);
      let mongoRecords = myObjects.documents;
      const mongoRecord = mongoRecords[0];
      const geoJSONobj = mongoRecord.geometry;
      Logger.log(geoJSONobj);
      const geojson_data = normalizeGeoJSON(geoJSONobj);
      Logger.log(geojson_data);
      var buffered = turf.buffer(geojson_data, (100 * 0.000189394), 'miles')
      // var buffered = turf.buffer(geojson_data, 100, {units: 'feet'});
      const geoJsonObj = buffered.geometry;
      const newGeoJsonObj = normalizeGeoJSONForGeojsonIO(geoJsonObj);
      const newGeojsonString = JSON.stringify(newGeoJsonObj)
      Logger.log(newGeojsonString);


      const encoded = encodeURIComponent(newGeojsonString);
      // const encoded = encodeURIComponent(newGeoJson);

      Logger.log(encoded);

      // http://geojson.io/#data=data:application/json,%7B%22type%22%3A%22LineString%22%2C%22coordinates%22%3A%5B%5B0%2C0%5D%2C%5B10%2C10%5D%5D%7D

      const longUrl = `http://geojson.io/#data=data:application/json,${encoded}`;

      // const longUrl = `https://geojson.io/next/#data=data:application/json,${encoded}`;
      // https://geojson.io/next/ -- note that this needs a url shortner - url too long


      Logger.log(longUrl);
      // const clickable = shortenUrl(longUrl);


      var ud = updateCell(sheet, myRow, 'RoadURL', longUrl);

    } catch (error) {
      Logger.log(error)
    }

    Logger.log("FIN")

  }

}


function createKMLFile(latLongPolygon, listing_id, title, description) {
  // Create KML content with title and description
  var kmlContent = '<?xml version="1.0" encoding="UTF-8"?>' +
    '<kml xmlns="http://www.opengis.net/kml/2.2">' +
    '<Document>' +
    '<Placemark>' +
    '<name>' + title + '</name>' +
    '<description>' + description + '</description>' +
    '<Polygon>' +
    '<outerBoundaryIs>' +
    '<LinearRing>' +
    '<coordinates>';

  // Add coordinates to KML content
  latLongPolygon.forEach(function (point) {
    kmlContent += point[1] + ',' + point[0] + ' ';
  });

  // Close KML content
  kmlContent += '</coordinates>' +
    '</LinearRing>' +
    '</outerBoundaryIs>' +
    '</Polygon>' +
    '</Placemark>' +
    '</Document>' +
    '</kml>';

  // Get the folder where you want to save the KML file
  var folder = DriveApp.getFolderById('1ssPWu5gWTQB-fW3YoOKagQxwITmc4AKy');

  // Save KML content to a file in the specified folder
  var fileName = `${listing_id}.kml`;

  var kmlFile = folder.createFile(fileName, kmlContent, 'application/vnd.google-earth.kml+xml');
  Logger.log('KML file created: ' + fileName + ' in folder: ' + folder.getName());
}


function aggregateMongoDBData(lon, lat, coll) {

  //  1GxuC9AAuc77xJklnS1PSHVKhZUt3QkcOaqdSYRqoeQVrSnf8jtGtLO3zGlNfm4T -- this is the api auth token and cannot be retrieved


  var url = "https://us-east-1.aws.data.mongodb-api.com/app/data-wygci/endpoint/data/v1/action/aggregate";
  var apiKey = "1GxuC9AAuc77xJklnS1PSHVKhZUt3QkcOaqdSYRqoeQVrSnf8jtGtLO3zGlNfm4T";
  var authToken = "1GxuC9AAuc77xJklnS1PSHVKhZUt3QkcOaqdSYRqoeQVrSnf8jtGtLO3zGlNfm4T";
  var headers = {
    "Content-Type": "application/json",
    "Access-Control-Request-Headers": "*",
    "api-key": apiKey,
    "Authorization": authToken,
  };


  var payload = {
    "pipeline": [
      {
        "$match": {
          "coordinate": { $exists: true, $ne: null }
        }
      },
      {
        "$set": {
          "distance": {
            "$expr": {
              "$function": {
                "body": function (lat1, lon1, lat2, lon2) {
                  // Haversine formula to calculate distance
                  var R = 6371; // km
                  var dLat = toRadians(lat2 - lat1);
                  var dLon = toRadians(lon2 - lon1);
                  var lat1Rad = toRadians(lat1);
                  var lat2Rad = toRadians(lat2);
                  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
                    Math.sin(dLon / 2) * Math.sin(dLon / 2) * Math.cos(lat1Rad) * Math.cos(lat2Rad);
                  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
                  var distance = R * c;
                  return distance;
                },
                "args": ["$coordinate.lat", "$coordinate.lon", lat, lon],
                "lang": "js"
              }
            }
          }
        }
      }
    ],
    "database": "mydata",
    "dataSource": "Cluster0",
    "collection": coll
  };

  var options = {
    method: "post",
    headers: headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  // const query = { list_date: { $gte: dt }, state: st };


  var response = UrlFetchApp.fetch(url, options);

  if (response.getResponseCode() === 200) {
    var result = response.getContentText();
    Logger.log(result);
    const myObjs = JSON.parse(result);
    Logger.log(myObjs.documents.length);

    return result;
  } else {
    Logger.log("Error: " + response.getResponseCode());
    Logger.log(response.getContentText());
    return false;
  }
}

function aggregateTest() {

  const myRow = getSelectedRowObject();

  // Logger.log(" comps : ");
  // // const myQuery = geoQuery(-87.214035,33.802722, 200);
  // const myQuery = geoQuery(Number(myRow.lon), Number(myRow.lat), Number(myRow.RADIUS), Number(myRow.MIN_ACRES), Number(myRow.MAX_ACRES));

  const res = aggregateMongoDBData(Number(myRow.lon), Number(myRow.lat), "wisconsinSold")
  Logger.log(res);



}

