// Inject custom menu `SEO Tools` when spreadsheet loads
function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('SEO Tools')
    .addItem('Google SERP', 'openDialog')
    .addToUi()
}

// Serve webpage when custom menu item `Google SERP` is clicked
function openDialog() {
  var html = HtmlService.createHtmlOutputFromFile('index')
  SpreadsheetApp.getUi()
    .showModalDialog(html, 'Google SERP')
}

// Handle form submission
async function submitQuery(params){
  try {
    let data = await search(params)
    let spreadsheet = SpreadsheetApp.getActiveSpreadsheet()

    // organic result sheet
    writeOrganicResults(data, spreadsheet)
  } catch (error) {
    // Handle error as desired
    Logger.log(error)
  }
}

// Object to querystring - https://gist.github.com/tanaikech/70503e0ea6998083fcb05c6d2a857107
String.prototype.addQuery = function(obj) {
  return this + Object.keys(obj).reduce(function(p, e, i) {
    return p + (i == 0 ? "?" : "&") +
      (Array.isArray(obj[e]) ? obj[e].reduce(function(str, f, j) {
        return str + e + "=" + encodeURIComponent(f) + (j != obj[e].length - 1 ? "&" : "")
      },"") : e + "=" + encodeURIComponent(obj[e]));
  },"");
}

// Call Google SERP on RapidAPI
function search(params) {
  return new Promise((resolve, reject) => {
    try {
      const X_RapidAPI_Key = params.rapid_api_key
      delete params.rapid_api_key

      let url = `https://google-search65.p.rapidapi.com/search`
      url = url.addQuery(params)

      let response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: {
          'X-RapidAPI-Key': X_RapidAPI_Key,
          'X-RapidAPI-Host': `google-search65.p.rapidapi.com`
        }
      })

      let data = response.getContentText()
      resolve(JSON.parse(data))
    } catch (error) {
      reject(error)
    }
  })
}

// Populate sheet with the `organic_result` data from RapidAPI
function writeOrganicResults(data, spreadsheet){
  Logger.log(`Writing data...📝`)
  let organic_results = data?.data?.organic_results

  if (organic_results.length < 1){
    return
  }

  let organicResultsSheet = spreadsheet.getSheetByName(`organic_results`)

  if (!organicResultsSheet) {
    spreadsheet.insertSheet(`organic_results`)
  }

  // Append search_metadata info at top of the result e.g device, location info, etc.
  writeSearchInfo(data, organicResultsSheet)

  // Append headers row
  organicResultsSheet.appendRow(Object.keys(organic_results[0]))

  // append the rest of the data
  organic_results.forEach((item) => {
    const keys = Object.keys(item)

    let rowData = keys.map((key) => {
      return item[key].toString()
    })

    organicResultsSheet.appendRow(rowData)
    Logger.log(`Row added to sheet! ✅`)
  })
}

function writeSearchInfo(data, organicResultsSheet){
  let search_query = data?.data?.search_query
  let headerContent = Object.keys(search_query)

  organicResultsSheet.appendRow(headerContent)
  
  let bodyContent = headerContent.map((item) => {
    return search_query[item]
  })

  organicResultsSheet.appendRow(bodyContent)
}
