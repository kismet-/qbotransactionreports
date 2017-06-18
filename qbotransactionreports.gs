var CONSUMER_KEY = '---';
var CONSUMER_SECRET = '---';
muteHttpExceptions=true

/**
 * Authorizes and makes a request to the QuickBooks API.
 */
function run() {
  var service = getService();
  if (service.hasAccess()) {
    var companyId = PropertiesService.getUserProperties()
        .getProperty('QuickBooks.companyId');
    var url = 'https://quickbooks.api.intuit.com/v3/company/' +
        companyId + '/reports/TransactionList?start_date=2017-01-01&end_date=2017-03-31&columns=tx_date,other_account,memo,subt_nat_amount&sort_by=tx_date'; // Edit date range and columns in URL
    var response = service.fetch(url, {
      headers: {
        'Accept': 'application/json',
      },
    })
    
    pullJSON(response.getContentText());
    
  } else {
    var authorizationUrl = service.authorize();
    Logger.log('Open the following URL and re-run the script: %s',
        authorizationUrl);
  }
}

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  var service = getService();
  service.reset();
}

/**
 * Configures the service.
 */
function getService() {
  return OAuth1.createService('QuickBooks')
      // Set the endpoint URLs.
      .setAccessTokenUrl('https://oauth.intuit.com/oauth/v1/get_access_token')
      .setRequestTokenUrl('https://oauth.intuit.com/oauth/v1/get_request_token')
      .setAuthorizationUrl('https://appcenter.intuit.com/Connect/Begin')

      // Set the consumer key and secret.
      .setConsumerKey(CONSUMER_KEY)
      .setConsumerSecret(CONSUMER_SECRET)

      // Set the name of the callback function in the script referenced
      // above that should be invoked to complete the OAuth flow.
      .setCallbackFunction('authCallback')

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties());
}

/**
 * Handles the OAuth callback.
 */
function authCallback(request) {
  var service = getService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    PropertiesService.getUserProperties()
        .setProperty('QuickBooks.companyId', request.parameter.realmId);
    return HtmlService.createHtmlOutput('Success!');
  } else {
    return HtmlService.createHtmlOutput('Denied');
  }
}

function CountValue( strText, reTargetString ){
    var intCount = 0;
    
    // Use replace to globally iterate over the matching
    // strings.
    strText.replace(
        reTargetString,
        
        // This function will get called for each match
        // of the regular expression.
        function(){
            intCount++; 
        }
    );
    
    // Return the updated count variable.
    return( intCount );
}


function pullJSON(json) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheet = ss.getActiveSheet();

  //var dataSet = json; 
  //var jsonFix =  JSON.stringify(dataSet);
//  var fi = jsonFix.replace(/.*?(?={"Row":)/, "");
//  fi = fi.replace(/.*?(?={"ColData":)/, "");
//  var ki = fi.replace(/"type":"Data"\},\{/g, "");
 

 // fi = JSON.parse(fi);
//  fi = ki.slice(0, -3);
 // for (i=0; i < 
// Logger.log(fi);
  
  
  //Logger.log(fi);
 // var i = 0;
  
//  fi = fi.replace(/ColData/g, function(m) { return "ColData" + (++i).toString(); });
//  Logger.log(fi);

 // fi = JSON.parse(fi);
  //Logger.log(JSON.stringify(fi));

  //Logger.log(JSON.parse(json)['Rows']);
  
  fi = JSON.parse(json);
 // Logger.log(fi.Rows.Row);
  var rows = fi.Rows.Row,
      data;
  

  tests = sheet.getRange(1,1,1,4)
  tests.setValues([["date", "memo", "account", "amount"]]);
  for (i = 0; i < rows.length; i++) {
   
   // Logger.log(rows[i].ColData[0]);
    
    hit = rows[i].ColData; 
    
    var meth2 = [];
    
    for (j=0; j< hit.length; j++){
      
     Logger.log(hit[j].value); 
      meth2.push(hit[j].value);
//     dataRange = sheet.getRange(i+2, j+1);
//     dataRange.setValue(hit[j].value);
     
    }
   tests = sheet.getRange(i+2,1,1,4)
  tests.setValues([meth2]);
    }
}
