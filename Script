function generateAnswer() {
  var questionCell = 'B7';
  var answerCell = 'G7';
  var apiKey = 'YOUR-API-KEY'; // Replace with your GPT-3.5 API key

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var question = sheet.getRange(questionCell).getValue();
  var infoSheetName = sheet.getRange('B11').getValue();
  var infoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(infoSheetName);
  var information = infoSheet.getDataRange().getValues();

  var prompt = 'Based on the information in the ' + infoSheetName + ' sheet, ' + question;

  var response = callGPTAPI(prompt, information, apiKey);
  var answer = response.choices[0].message.content.trim();

  sheet.getRange(answerCell).setValue(answer);
}

function callGPTAPI(prompt, information, apiKey) {
  var apiUrl = 'https://api.openai.com/v1/chat/completions';
  var headers = {
    'Authorization': 'Bearer ' + apiKey,
    'Content-Type': 'application/json'
  };
  var data = {
    'messages': [
      {
        'role': 'system',
        'content': 'You are a user.'
      },
      {
        'role': 'user',
        'content': prompt
      }
    ],
    'max_tokens': 100,
    'temperature': 0.5,
    'n': 1,
    'model': 'gpt-3.5-turbo'
  };

  // Append the information from the sheet to the user's prompt
  information.forEach(function(row) {
    var rowText = row.join(' ');
    data.messages.push({
      'role': 'system',
      'content': rowText
    });
  });

  var options = {
    'method': 'post',
    'headers': headers,
    'payload': JSON.stringify(data)
  };

  var response = UrlFetchApp.fetch(apiUrl, options);
  return JSON.parse(response.getContentText());
}
