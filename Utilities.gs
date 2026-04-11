function extractSubstring(corpus, startString, endString) {

  return corpus.substring(
    corpus.indexOf(startString) + 1,
    corpus.lastIndexOf(endString)
  );

}

function createCountySubdivisionSummaries() {

  // - NOTE TO SEND TO netlify update Bucket function it must be an array with a property_id 

  // const counties = ["Oconto", "Lincoln", "Taylor", "Florence", "Rusk", "Barron", "Forest", "Sawyer", "Polk", "Price", "Chippewa", "Clark", "Langlade", "Ashland", "Iron"];
  // const counties = ["Oconto", "Lincoln", "Taylor"];
  // const countyJson = JSON.stringify(counties);
  // Logger.log(countyJson);

  const ui = SpreadsheetApp.getUi();

  // First prompt: list of counties (comma-separated)
  const countyResp = ui.prompt(
    'Counties',
    'Enter county names separated by commas (e.g. Oconto, Lincoln, Taylor):',
    ui.ButtonSet.OK_CANCEL
  );

  if (countyResp.getSelectedButton() !== ui.Button.OK) return; // user cancelled

  const counties = countyResp
    .getResponseText()
    .split(',')
    .map(c => c.trim())
    .filter(c => c); // remove empty entries

  // Second prompt: single state
  const stateResp = ui.prompt(
    'State',
    'Enter the state (e.g. WI):',
    ui.ButtonSet.OK_CANCEL
  );

  if (stateResp.getSelectedButton() !== ui.Button.OK) return; // user cancelled

  const state = stateResp.getResponseText().trim();

  // Use counties[] and state here
  Logger.log(counties);
  Logger.log(state);




  let url = 'https://nimble-dieffenbachia-92e8e2.netlify.app/.netlify/functions/openRouter2';

  Logger.log(url);

  // const options = {
  //   "method": "post",
  //   "contentType": "application/json",
  //   "payload": countyJson
  // };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(counties),
    muteHttpExceptions: true
  };


  Logger.log("options");

  Logger.log(options);
  // return;


  // Make the API request
  const response = UrlFetchApp.fetch(url, options);
  var responseObj = JSON.parse(response.getContentText());
  const resultsArr = JSON.parse(responseObj.results);

  Logger.log(resultsArr);

  // https://docs.google.com/document/d/1RWNp0xpmcxHm0HZSflTGpwPDGaXMbuM-WVqL_uN8x_A/edit?tab=t.0

  // Open the Google Doc by ID
  const docId = '1RWNp0xpmcxHm0HZSflTGpwPDGaXMbuM-WVqL_uN8x_A'; // Replace with your actual document ID
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();

  // Loop through rows and append each to the document
  for (var i = 0; i < resultsArr.length; i++) {
    const results = resultsArr[i];
    body.appendParagraph(results.county);
    body.appendParagraph(results.result);


  }



}


/**
 * Removes characters from the beginning of a string up to the first occurrence of a marker.
 * Returns the modified string. If the marker is not found, returns the original string.
 *
 * @param {string} originalMessage The input string.
 * @param {string} marker The string to search for (e.g., "Hi").
 * @return {string} The string starting from the marker, or the original string if the marker is not found.
 */
function removePrefixUntil(originalMessage, marker) {
  // Find the index of the first occurrence of the marker
  let startIndex = originalMessage.indexOf(marker);

  // Check if the marker was found
  if (startIndex !== -1) {
    // If found, use slice() to get the part of the string from that index to the end
    return originalMessage.slice(startIndex);
  } else {
    // If the marker was not found, return the original message
    Logger.log(`Marker "${marker}" not found in the message.`);
    return originalMessage; // Or return "" if you prefer an empty string
  }
}

function getFunctionByFlag(flag) {
  // Return an arrow function based on the input flag
  switch (flag.toLowerCase()) {
    case 'option1':
      return (location, min, max) =>
        `https://www.realtor.com/realestateandhomes-search/${location}/type-land/lot-sqft-${min}-${max}/sby-6?view=map`;

    case 'option2':
      return (location, min, max) =>
        `https://www.another-site.com/search/${location}/land-area-${min}-${max}`;

    case 'option3':
      return (location, min, max) =>
        `https://www.some-other-site.com/properties/${location}/size-${min}-to-${max}`;

    default:
      return () => 'Invalid option. Please provide a valid flag like option1, option2, or option3.';
  }
}

// Example usage
const flag = 'option1'; // Choose which function to return
const selectedFunction = getFunctionByFlag(flag); // Get the function, but don't execute it yet

// Later, execute the function
const url = selectedFunction('City1', 1000, 5000);
Logger.log(url);
// Outputs: 'https://www.realtor.com/realestateandhomes-search/City1/type-land/lot-sqft-1000-5000/sby-6?view=map'



function createSMSgatewayEmail() {

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  const myRow = getSelectedRowObject();
  console.log(myRow);
  const phoneNumber = myRow.PHONE;
  const formatted = formatPhoneNumber(phoneNumber);
  const out = checkUserPhone(formatted);
  console.log(out);
  const sms_email = getGateway(out);
  // sms_email=7159380056@cellcom.quiktxt.com
  var ud = updateCell(sheet, myRow, 'EmailToTexT', sms_email);

}

function formatPhoneNumber(phoneNumber) {
  // Remove non-numeric characters using a regex
  var cleaned = phoneNumber.replace(/[^\d]/g, '');

  // Add the country code (1) to the beginning of the number
  var formattedNumber = '1' + cleaned;

  return formattedNumber;
}

function getGateway(provider) {

  // { message: 'Phone is valid.',
  //   success: true,
  //   formatted: '+16174551705',
  //   local_format: '(617) 455-1705',
  //   valid: true,
  //   fraud_score: 0,
  //   recent_abuse: false,
  //   VOIP: false,
  //   prepaid: false,
  //   risky: false,
  //   active: true,
  //   carrier: 'Verizon Wireless',
  //   line_type: 'Wireless',
  //   country: 'US',
  //   city: 'EAST BOSTON',
  //   zip_code: '02128',
  //   region: 'MA',
  //   dialing_code: 1,
  //   active_status: 'N/A',
  //   sms_domain: 'vtext.com',
  //   associated_email_addresses: { status: 'Enterprise Plus or higher required.', emails: [] },
  //   user_activity: 'Enterprise L4+ required.',
  //   mnc: '004',
  //   mcc: '310',
  //   leaked: true,
  //   spammer: false,
  //   request_id: 'Qxi45C4umX',
  //   name: '',
  //   timezone: 'America/New_York',
  //   do_not_call: true,
  //   accurate_country_code: false,
  //   sms_email: '6174551705@vtext.com' }

  // https://en.wikipedia.org/wiki/SMS_gateway#Email_clients

  const carrierGateways = [
    {
      shortName: "Alltel",
      carrier: "Alltel",
      sms: "sms.alltelwireless.com",
      mms: "mms.alltelwireless.com"
    },
    {
      shortName: "AT&T",
      carrier: "AT&T",
      sms: "txt.att.net",
      mms: "mms.att.net"
    },
    {
      shortName: "Boost",
      carrier: "Boost Mobile",
      sms: "sms.myboostmobile.com",
      mms: "myboostmobile.com"
    },
    {
      shortName: "Consumer",
      carrier: "Consumer Cellular",
      sms: "mailmymobile.net",
      mms: "mailmymobile.net"
    },
    {
      shortName: "Cricket",
      carrier: "Cricket Wireless",
      sms: "mms.cricketwireless.net",
      mms: "mms.cricketwireless.net"
    },
    {
      shortName: "Google",
      carrier: "Google Fi Wireless",
      sms: "msg.fi.google.com",
      mms: "msg.fi.google.com"
    },
    {
      shortName: "MetroPCS",
      carrier: "MetroPCS",
      sms: "mymetropcs.com",
      mms: "mymetropcs.com"
    },
    {
      shortName: "Republic",
      carrier: "Republic Wireless",
      sms: "text.republicwireless.com",
      mms: ""
    },
    {
      shortName: "Sprint",
      carrier: "Sprint",
      sms: "messaging.sprintpcs.com",
      mms: "pm.sprint.com"
    },
    {
      shortName: "T-Mobile",
      carrier: "T-Mobile",
      sms: "tmomail.net",
      mms: "tmomail.net"
    },
    {
      shortName: "Ting",
      carrier: "Ting",
      sms: "message.ting.com",
      mms: ""
    },
    {
      shortName: "U.S.",
      carrier: "U.S. Cellular",
      sms: "email.uscc.net",
      mms: "mms.uscc.net"
    },
    {
      shortName: "Verizon",
      carrier: "Verizon Wireless",
      sms: "vtext.com",
      mms: "vzwpix.com"
    },
    {
      shortName: "Virgin",
      carrier: "Virgin Mobile",
      sms: "vmobl.com",
      mms: "vmpix.com"
    },
    {
      shortName: "XFinity",
      carrier: "XFinity Mobile",
      sms: "vtext.com",
      mms: "mypixmessages.com"
    }
  ];

  let smsEmail = false;
  for (var i = 0; i < carrierGateways.length; i++) {
    const carrierGateway = carrierGateways[i];
    if (provider.carrier.includes(carrierGateway.shortName)) {
      smsEmail = provider.sms_email.split("@")[0] + "@" + carrierGateway.sms;
    }

  }
  return smsEmail;
}



function getRealtorcomAgentURL() {

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  const myRow = getSelectedRowObject();
  console.log(myRow);
  const state = myRow.state;
  const county = myRow.county;
  const row = lookupZipCodesByCountyState(county, state);
  myURL = `https://www.realtor.com/realestateagents/${row.zip}`
  var ud = updateCell(sheet, myRow, 'COMP_AVG_DISTANCE', myURL);

}

function lookupZipCodesByCountyState(county, state) {

  var spreadsheet = SpreadsheetApp.openById('1xuZhApjf-mq7hlMkIVlUo9eqA9TT8o5bYybF_m17tZ4');
  console.log(spreadsheet.getName());
  const sheet = spreadsheet.getSheetByName('zip_code_database')
  console.log(sheet.getName());

  const json = sheet2Json(sheet);

  let myRow = false;
  for (var i = 0; i < json.length; i++) {
    const row = json[i];
    if (row.state == state && row.county.includes(county)) {
      if (!myRow) {
        myRow = row;
      } else if (row.irs_estimated_population > myRow.irs_estimated_population) {
        myRow = row;


      }
    }
  }
  console.log(myRow);
  return myRow

}

function checkUserPhoneTest() {
  const result = checkUserPhone("17153604366");
  console.log(result);
}


function checkUserPhone(phone) {
  // https://www.ipqualityscore.com/documentation/phone-number-validation-api/overview
  // phone should b 1dddddddddd (only numbers)
  var key = '5AALTL2ikjg4qz8UP8aHmKXXTkfuVMKj';
  var url = "https://www.ipqualityscore.com/api/json/phone/" + key + "/" + encodeURIComponent(phone);

  var result = getIPQURL(url);
  if (result !== null) {
    return result;
  } else {
    throw new Error('No response received from IPQualityScore API.');
  }
}

function getIPQURL(url) {
  try {
    var response = UrlFetchApp.fetch(url);
    return JSON.parse(response.getContentText());
  } catch (error) {
    Logger.log('Error fetching URL: ' + error.message);
    return null;
  }
}



// function contactLocalAgents() {

//   const myRow = getSelectedRowObject();
//   console.log(myRow);
//   const address = myRow.address;
//   const state = myRow.state;
//   const county = myRow.county;
//   const zips = lookupZipCodesByCountyState(county, state)
//   console.log(zips);

// }



function IEvalDate() {

  const myRow = getSelectedRowObject();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var today = Utilities.formatDate(new Date(), "GMT+1", "MM/dd/yyyy")
  var ud = updateCell(sheet, myRow, 'InitialEvaluation', today);

}


function getRangeSTR() {

  // var selection = SpreadsheetApp.getSelection();
  // var activerow = selection.getCurrentCell().getRow();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var strSheet = spreadsheet.getSheetByName('STR');
  var sheet = spreadsheet.getActiveSheet();


  const json = sheet2Json(strSheet);
  // const myRow = json[activerow - 2];

  // Logger.log(myRow.lon + ", " + myRow.lat);
  // const longUrl = createFemaUrl(myRow.lat, myRow.lon);
  // // const respJson = UrlFetchApp.fetch(longUrl);
  // // const finalUrl = constructMapServerURL(respJson)
  // // Logger.log(finalUrl);

  // const clickable = shortenUrl(longUrl);


  // var ud = updateCell(sheet, myRow, 'FLOODPLAIN', clickable);


  const myRow = getSelectedRowObject();
  console.log(myRow);
  const min = myRow.MIN_ACRES;
  const max = myRow.MAX_ACRES;
  const mid = (max + min) / 2;
  const county = myRow.county;
  console.log(mid);

  let strRow = false;
  for (var i = 0; i < json.length; i++) {
    console.log(json[i]);
    const rowObj = json[i];
    if (rowObj.COUNTY == county) {
      strRow = rowObj;
    }

  }

  console.log(strRow);

  if (!strRow) { throw new Error("No county match!!."); return; }

  let strValue = false;
  const keys = Object.keys(strRow);
  console.log(keys);

  for (var j = 0; j < keys.length; j++) {
    const key = keys[j];
    if (key.includes("--")) {
      const pieces = key.split("--");
      if (mid >= pieces[0] && mid <= pieces[1] && pieces[2] == "STR") {
        strValue = strRow[key];
      }
    }
  }


  var ud = updateCell(sheet, myRow, 'STR', strValue);







}




function addNote(text = "", column = "NOTES") {

  const myRow = getSelectedRowObject();
  const cr = getRowCellRangeByColumnName(myRow, column);
  console.log(cr);

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get the sheet by name
  var sheet = spreadsheet.getSheetByName('Sheet1');

  // Get the specific cell (e.g., A1)
  var cell = sheet.getRange(cr);

  // Retrieve the existing note
  var currentNote = cell.getNote();

  // Define the new bullet point to add
  var newBulletPoint = "• " + text;

  // Append the new bullet point to the existing note
  var updatedNote = currentNote + "\n" + newBulletPoint;

  // Update the cell with the new note
  cell.setNote(updatedNote);

}




function extractAPN() {

  const myRow = getSelectedRowObject();
  console.log(myRow);
  const myUrl = myRow.AAlink;
  // const myUrl = "https://www.realtor.com/realestateandhomes-detail/55470-State-Highway-171_Seneca_WI_00000_M91305-68892?from=srp"

  var url = "https://realio.onrender.com/detailPageData?val=" + myUrl;

  var options = {
    "method": "GET",
    "headers": {
      "Accept": "application/json"
    }
  };

  var response = UrlFetchApp.fetch(url, options);

  const mydata = JSON.parse(response);
  //  console.log(JSON.stringify(response))


  // console.log(mydata.data.props.pageProps);
  // console.log(mydata.data.props.pageProps.initialReduxState.propertyDetails.description);
  console.log(mydata.data.props.pageProps.initialReduxState.propertyDetails);

  const prompt = "Extract 'Parcel Number' and return the Number only from the following:";

  const resp = makeHuggingfacePostRequest(JSON.stringify(mydata.data.props.pageProps.initialReduxState.propertyDetails), prompt)


  const output = JSON.parse(resp)
  const parcelNumber = output.response;
  const stripped = parcelNumber.replace(/(<([^>]+)>)/ig, '');


  Logger.log("Parcel Number: " + stripped);

  var sheet = SpreadsheetApp.getActiveSheet();

  var ud = updateCell(sheet, myRow, 'APN2', stripped);

}

function makeHuggingfacePostRequest(contextstring, prompt) {
  var url = "https://kc11-mixtral-46-7b-fastapi.hf.space/generate/";

  var payload = {
    // "prompt": "What is the capital of France?",
    // "prompt": `Extract 'Parcel Number' from the context`,
    "prompt": `${prompt} ${contextstring} `,
    "history": [],
    "system_prompt": "You are a very powerful AI assistant.",
    // "context": contextstring
  };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };

  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText());
  return response;
}


function createTnmUrl() {

  const myRow = getSelectedRowObject();
  const coords = JSON.parse(myRow.BB);


  var baseUrl = "https://tnmaccess.nationalmap.gov/api/v1/products";
  var bbox = coords.west + "," + coords.south + "," + coords.east + "," + coords.north;
  var datasets = "National%20Elevation%20Dataset%20(NED)%201/3%20arc-second%20Current";

  var url = baseUrl + "?bbox=" + bbox + "&datasets=" + datasets;
  var options = {
    "method": "GET",
    "headers": {
      "Accept": "application/json"
    }
  };

  var response = UrlFetchApp.fetch(url, options);
  var jsonData = JSON.parse(response.getContentText());

  var downloadURL = jsonData.items[0].downloadURL;


  var sheet = SpreadsheetApp.getActiveSheet();

  var ud = updateCell(sheet, myRow, 'tnmDEM', downloadURL);
}



function buildBB() {

  const myRow = getSelectedRowObject();
  const wkt = myRow.GEOM;
  const polygon = parseWKTToPolygon(wkt);
  Logger.log(polygon);

  var reversedPolygon = polygon.map(subArr => subArr.reverse());

  Logger.log(reversedPolygon); // Output: [[b, a], [d, c]]
  const bb = getBoundingBoxFromPolygonCoordinates(reversedPolygon);
  Logger.log(bb);

  var sheet = SpreadsheetApp.getActiveSheet();

  var ud = updateCell(sheet, myRow, 'BB', JSON.stringify(bb));


}

function opentopographyDEM(bb) {

  var apiUrl = "https://portal.opentopography.org/API/usgsdem";
  var url = "https://portal.opentopography.org/API/usgsdem?datasetName=USGS10m&south=40.234&north=40.288&west=-105.673&east=-105.583&outputFormat=GTiff&API_Key=e8d224c1ac5b43c481259b709ccf9d52";

  // var params = {
  //   "datasetName": "USGS30m",
  //   "south": bb.south,
  //   "north": bb.north,
  //   "west":  bb.west,
  //   "east": bb.east,
  //   "outputFormat": "GTiff",
  //   "API_Key": "e8d224c1ac5b43c481259b709ccf9d52"
  // };
  // Logger.log(params);

  // var options = {
  //   "method": "GET",
  //   // "headers": {
  //   //   "Authorization": "Bearer " + params.API_Key
  //   // },
  //   "params": params
  // };

  var response = UrlFetchApp.fetch(url);
  var demData = response.getContentText();

  // Display the output
  var output = "DEM Data:\n" + demData;
  Logger.log(output);
  return output;
}



function getViewportInfo() {
  var url = "https://maps.googleapis.com/$rpc/google.internal.maps.mapsjs.v1.MapsJsInternalService/GetViewportInfo";
  var headers = {
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.9",
    "Content-Type": "application/json+protobuf",
    "DNT": "1",
    "Origin": "https://contourmapcreator.urgr8.ch",
    "Priority": "u=1, i",
    "Referer": "https://contourmapcreator.urgr8.ch/",
    "Sec-Ch-Ua": "\"Google Chrome\";v=\"125\", \"Chromium\";v=\"125\", \"Not.A/Brand\";v=\"24\"",
    "Sec-Ch-Ua-Mobile": "?0",
    "Sec-Ch-Ua-Platform": "\"Chrome OS\"",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "cross-site",
    "User-Agent": "Mozilla/5.0 (X11; CrOS x86_64 14541.0.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "X-Client-Data": "CIe2yQEIo7bJAQipncoBCPLiygEIlKHLAQiIoM0BCIChzgEIpqLOARj0yc0BGOuNpRc=",
    "X-Goog-Api-Key": "AIzaSyAaq--1Pj7CYE_G-llqg94Il3Lxzk1c3Ao",
    "X-Goog-Maps-Api-Salt": "Sol7QGHsd",
    "X-Goog-Maps-Api-Signature": "27943",
    "X-Goog-Maps-Channel-Id": "",
    "X-Goog-Maps-Client-Id": "",
    "X-User-Agent": "grpc-web-javascript/0.1"
  };

  var data = [[[46.140763172727304, 89.23585783155977], [46.168410529661635, 89.28847699475736]], 16, null, "en-US", 3, "m@696000000", 0, 0, null, null, null, 2];

  var options = {
    "method": "POST",
    "headers": headers,
    "payload": JSON.stringify(data)
  };

  var response = UrlFetchApp.fetch(url, options);
  var result = JSON.parse(response.getContentText());
  Logger.log(result);



  // Alternatively, create the file in the root folder:
  var file = DriveApp.createFile("Contour.json", JSON.stringify(result, null, 2));

  Logger.log("Result saved to Drive: " + file.getUrl());

}


function getHTMLWithScrapingant2(buildURL) {
  var encodedURL = encodeURIComponent(buildURL);
  var apiUrl = "https://api.scrapingant.com/v2/general";
  var options = {
    "method": "GET",
    "headers": {
      "useQueryString": true
    },
    "params": {
      "url": encodedURL,
      "x-api-key": "3c59fe0e311a474694cd2849f594f135",
      "return_page_source": true
    }
  };

  try {
    var response = UrlFetchApp.fetch(apiUrl, options);
    if (response.getResponseCode() == 200) {
      var body = response.getContentText();
      Logger.log(body);
      return body;
    } else {
      Logger.log("Request failed with status " + response.getResponseCode());
      return false;
    }
  } catch (error) {
    Logger.log("An error occurred: " + error);
  }
}



function shortenUrl(longUrl) {
  var apiEndpoint = "https://is.gd/api.php?longurl=" + encodeURIComponent(longUrl);
  var response = UrlFetchApp.fetch(apiEndpoint);
  var shortUrl = response.getContentText();
  return shortUrl;

}

function loadFileContents(fileName) {
  var folderId = '1ssPWu5gWTQB-fW3YoOKagQxwITmc4AKy'; // Replace with the ID of the folder
  // var fileName ='myFile.txt'; // Replace with the name of the file
  var folder = DriveApp.getFolderById(folderId);
  var file = folder.getFilesByName(fileName).next();
  var contents = file.getBlob().getDataAsString();
  Logger.log(contents); // Log the file contents as a string
  return contents;
}

function encodeKml(kmlStr) {
  var encoded = Utilities.base64EncodeWebSafe(kmlStr);
  return encoded;
}

function getCenterPointFromBoundingBox(boundingBox) {
  var latCenter = (boundingBox.north + boundingBox.south) / 2;
  var lonCenter = (boundingBox.east + boundingBox.west) / 2;

  return [lonCenter, latCenter];
}

function getBoundingBoxFromPolygonCoordinates(coordinates) {
  var latMin = Infinity;
  var latMax = -Infinity;
  var lonMin = Infinity;
  var lonMax = -Infinity;

  for (var i = 0; i < coordinates.length; i++) {
    var lat = coordinates[i][1];
    var lon = coordinates[i][0];

    if (lat < latMin) latMin = lat;
    if (lat > latMax) latMax = lat;
    if (lon < lonMin) lonMin = lon;
    if (lon > lonMax) lonMax = lon;
  }

  return {
    north: latMax,
    south: latMin,
    east: lonMax,
    west: lonMin
  };
}

function parseWKTToPolygon(wkt) {
  // Extract the coordinates from the WKT string
  var coordinates = wkt.match(/-?\d+\.\d+/g);

  // Convert the flat array of coordinates to an array of arrays
  var polygon = [];
  for (var i = 0; i < coordinates.length; i += 2) {
    var lat = parseFloat(coordinates[i + 1]);
    var lon = parseFloat(coordinates[i]);
    polygon.push([lat, lon]);
  }

  Logger.log(polygon);
  return polygon;
}

function constructMapServerURL(jsonInput) {
  // Parse the input JSON object
  var input = JSON.parse(jsonInput);

  // Extract the extent values from the first candidate in the JSON object
  var extent = input.candidates[0].extent;
  var xmin = extent.xmin;
  var ymin = extent.ymin;
  var xmax = extent.xmax;
  var ymax = extent.ymax;

  // Construct the bbox parameter
  var bbox = xmin + ',' + ymin + ',' + xmax + ',' + ymax;

  // Define the base URL and other parameters
  var baseURL = 'https://fwsprimary.wim.usgs.gov/server/rest/services/Test/Wetlands_gdb_split/MapServer/export';
  var params = {
    dpi: 96,
    transparent: true,
    format: 'png8',
    bbox: bbox,
    bboxSR: 3857,
    imageSR: 3857,
    size: '835,817',
    f: 'image'
  };

  // Construct the final URL
  var url = baseURL + '?' + Object.keys(params).map(function (key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
  }).join('&');

  // Return the constructed URL
  return url;
}

function constructGeocodeURL(latitude, longitude) {
  // Base URL for the geocode service
  var baseURL = 'https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/findAddressCandidates';

  // Define the parameters
  var params = {
    SingleLine: longitude + ", " + latitude,
    f: 'json',
    outSR: '{"wkid":3857}',
    maxLocations: 25
  };

  // Construct the final URL
  var url = baseURL + '?' + Object.keys(params).map(function (key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
  }).join('&');

  // Return the constructed URL
  Logger.log("constructGeocodeURL");
  Logger.log(url);

  return url;
}

// // Example usage
// var latitude = 45.736969;
// var longitude = -91.64949;
// Logger.log(constructGeocodeURL(latitude, longitude));


function getDetail(myLink) {


  // var selection = SpreadsheetApp.getSelection();
  // var activerow = selection.getCurrentCell().getRow();
  // var sheet = SpreadsheetApp.getActiveSheet();

  // const json = sheet2Json(sheet);
  // const myRow = json[activerow - 2];
  // // const myRow = getSelectedRowObject();
  // const myLink = myRow.AAlink;
  // // const encodedLink = encodeURI(myLink);
  // // Logger.log(encodedLink);

  // const url = 'https://realio1-c51a04e6b1da.herokuapp.com/getAgent';
  const url = "https://realio.onrender.com/detailPageData";
  Logger.log(url);

  var headers = {
    'sec-ch-ua': '"Chromium";v="118", "Google Chrome";v="118", "Not=A?Brand";v="99"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Chrome OS"',
    'Referer': 'https://realio.onrender.com/',
    'Referrer-Policy': 'strict-origin-when-cross-origin',
    'Content-Type': 'application/x-www-form-urlencoded', // Add this line for POST requests
  };

  // Create the payload for the POST request
  var payload = {
    'val': myLink,
  };

  var options = {
    'method': 'post', // Change the method to "post"
    'headers': headers,
    'payload': payload, // Include the payload in the request body
    'muteHttpExceptions': true
  };

  var response = UrlFetchApp.fetch(url, options);
  var content = response.getContentText();

  // Handle the response content as needed.
  Logger.log(content);

  const myObj = JSON.parse(content);
  // const agentDataObj = myObj.message[0];
  // const nameStr = agentDataObj.name;
  // const emailStr = agentDataObj.emails.join(',');
  // const phoneStr = agentDataObj.phones.join(",");

  // // myRow.AgentName = nameStr;
  // // myRow.AgentEmail = emailStr;
  // // myRow.AgentPhone = phoneStr;

  // var AN = updateCell(sheet, myRow, 'AgentName', nameStr);
  // var AE = updateCell(sheet, myRow, 'AgentEmail', emailStr);
  // var AP = updateCell(sheet, myRow, 'AgentPhone', phoneStr);

  return myObj;
}


function getSubdivComps() {

  // {AAlink=https://www.realtor.com/realestateandhomes-detail/Alma-Rd_Jasper_AL_35501_M98724-43725?from=srp, AgentEmail=, address=Alma Rd, list_date=2024-03-17T23:25:39.000000Z, ACRES/PIECE=, =, flags=is_new_listing, RADIUS=20.0, AgentName=, state=Walker, county=2024-03-18T00:18:19.665Z, APN=, ppa=4000.0, RANGE_PPA=, lot_acres=41.0, updatedAt=al, listing_id=2.965226047E9, price=164000.0, NOTES=, GEOM=POLYGON((-111.179167643 32.5349351874,-111.179166834 32.5358422453,-111.180237626 32.5358413441,-111.180238435 32.5349342937,-111.179167643 32.5349351874)), PIECES=, COMP_AVG_DISTANCE=, relativeRow=4.0, APN2=, MIN_ACRES=5.0, lon=-87.214035, AgentPhone=, MAX_ACRES=10.0, absoluteRow=5.0, lat=33.802722}

  const myRow = getSelectedRowObject();

  // Logger.log(myRow);
  Logger.log(" comps : ");
  // const myQuery = geoQuery(-87.214035,33.802722, 200);
  const myQuery = geoQuery(Number(myRow.lon), Number(myRow.lat), Number(myRow.RADIUS), Number(myRow.MIN_ACRES), Number(myRow.MAX_ACRES));
  // const myQuery = { "location": { "$nearSphere": { "$geometry": { "type": "Point", "coordinates": [-87.214035, 33.802722] }, "$maxDistance": 321868 } }, "acres": { "$gte": 5, "$lte": 10 } }

  console.log(JSON.stringify(myQuery));


  const myJSON = fetchMongoDBDataAPI(myQuery, "alabamaLandcomComps");
  Logger.log(myJSON);
  const myObjects = JSON.parse(myJSON);

  var sheetName = "landcomanalysis"; // replace with the name of your sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  Logger.log("OUT")

  Logger.log(sheet.getLastRow())
  // if (sheet.getLastRow() > 2) {
  sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();
  // }

  // var jsonText = body.getText();
  // var json = JSON.parse(jsonText);
  const propertyResults = myObjects.documents;
  // Do something with the JSON data, e.g., log it to the console
  Logger.log(propertyResults);

  for (var i = 0; i < propertyResults.length; i++) {
    const property = propertyResults[i];
    Logger.log(propertyResults[i]);

    const newRow = addRowFromObject(property, sheet);
    sheet.appendRow(newRow);

    Logger.log(property);


  }
}


function createNewSpreadsheetFromTemplate() {

  const myRow = getSelectedRowObject();
  const detailObj = getDetail(myRow.AAlink);

  Logger.log(detailObj);
  return;
  Logger.log(" comps : ");
  var templateSpreadsheetId = '1fbu5Usvnlk8zc2h4Z_MIu8WobBl8P3yVVPuirsM0RkQ'; // Replace with the ID of your template spreadsheet
  var newSpreadsheetName = 'New Spreadsheet' + new Date().getTime(); // Create a unique name for the new spreadsheet

  // Create a new spreadsheet as a copy of the template
  var newSpreadsheet = DriveApp.getFileById(templateSpreadsheetId).makeCopy(newSpreadsheetName);

  // Get the ID of the new spreadsheet
  var newSpreadsheetId = newSpreadsheet.getId();

  // Open the new spreadsheet
  var newSpreadsheet = SpreadsheetApp.openById(newSpreadsheetId);

  // Optional: You can also set the new spreadsheet's properties, such as the folder or permissions
  // newSpreadsheet.setFolderById('FOLDER_ID'); // Replace with the ID of the folder where you want to save the new spreadsheet
  // newSpreadsheet.addEditor('EDITOR_EMAIL'); // Replace with the email address of the user you want to add as an editor
  var sheet = newSpreadsheet.getSheetByName('Sheet1');

  sheet.getRange('A136').setValue(myRow);


  Logger.log('New spreadsheet created: %s', newSpreadsheet.getUrl());
}