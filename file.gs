const sheet = SpreadsheetApp.getActiveSpreadsheet();
const mySheet = sheet.getActiveSheet()
let grossPrice = []

let threads = GmailApp.search("from:YOUR_MAIL");

  let messages = threads[0].getMessages();
  let message = messages[messages.length - 1];
  let attachment = message.getAttachments()[0];
  attachment.setContentTypeFromExtension()
  
  let data = [];
  
  if (attachment.getContentType() == MimeType.CSV) {
    data = Utilities.parseCsv(attachment.getDataAsString(), ";");
  } else if (attachment.getContentType() == MimeType.MICROSOFT_EXCEL || attachment.getContentType() == MimeType.MICROSOFT_EXCEL_LEGACY) {
    let tempFile = Drive.Files.insert({title: "temp", mimeType: MimeType.GOOGLE_SHEETS}, attachment).id;
    data = SpreadsheetApp.openById(tempFile).getSheets()[0].getDataRange().getValues();
    Drive.Files.trash(tempFile);
  }
  
function main() {
  
  if (data.length > 0) {
    mySheet.clearContents().clearFormats();
    mySheet.getRange(mySheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
    
    calculateGrossPrice();
    formatTaxRate();
    fillGrossPriceColumn();   
    sendEmails();
  }
}


function formatTaxRate() {
  for (let i = 0; i < data.length - 1; i++) {
      let cells = mySheet.getRange(i + 2, 5).getValue();
      cells = cells * 100
      mySheet.getRange(i + 2, 5).setValue(cells + '%')
    }
}

function calculateGrossPrice() {
  for (let i = 0; i < data.length - 1; i++) {
    let netPrice = mySheet.getRange(i + 2, 4).getValue();
    let tax = mySheet.getRange(i + 2, 5).getValue();
    grossPrice.push(netPrice * (tax + 1))
  }
}

function fillGrossPriceColumn() {
  mySheet.getRange("H1").setValue('car_price_gross')
    for (let i = 0; i < data.length - 1; i++) {
      mySheet.getRange(i + 2, 8).setValue(grossPrice[i])
    } 
}

function sendEmails() {
  for (let i = 0; i < data.length - 1; i++) {
    let stockId = mySheet.getRange(i + 2, 1).getValue();
    let carBrand = mySheet.getRange(i + 2, 2).getValue();
    let carType = mySheet.getRange(i + 2, 3).getValue();
    let carPrice = mySheet.getRange(i + 2, 8).getValue();
    let email = mySheet.getRange(i + 2, 7).getValue();
    
    let subject = "Confirmation of your purchase"
    let message = "Congratulations, you have bought the following car:" + " " + stockId + " " + carBrand + " " + carType + " " + "for the price" + " " + carPrice + " " + "euro";
    
    
    MailApp.sendEmail(email, subject, message);
    Utilities.sleep(3 * 1000)
  }
  
}

