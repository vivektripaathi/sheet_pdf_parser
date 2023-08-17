const { GoogleSpreadsheet } = require('google-spreadsheet');
const fs = require("fs");
const { PdfReader } = require("pdfreader");
const https = require('https');
const creds = require('./clientSecret.json');

var rows, url, r = 1, flag = false, flag2 = true;

async function accessSpreadsheet(){
    const doc  = new GoogleSpreadsheet('1HulSPTgBVo_--6BFk2RcTx296JXIHAKaSU_mOu_KIiI');
    await doc.useServiceAccountAuth(creds);
    await doc.loadInfo();
    const sheet = doc.sheetsByIndex[0];
    rows = await sheet.getRows();
    flow();
}
accessSpreadsheet();

function flow(){
    if(r<rows.length){
        const row = rows[r];
        url = row.InvoiceDownloadLink;
        download(url);
    }
}

async function download(url){
    https.get(url, async (res) => {
        const path = `${__dirname}/invoice.pdf`; 
        const filePath = fs.createWriteStream(path);
        res.pipe(filePath);
        filePath.on('finish',async () => {
            filePath.close();
            pdfExt();
        });
    });
    return 0;
}

async function pdfExt() {
    var flag = false;
    var temp = "";
    var k = "";
    var res = {};
    fs.readFile("invoice.pdf", (err, pdfBuffer) => {
      // pdfBuffer contains the file content
      new PdfReader().parseBuffer(pdfBuffer, (err, item) => {
        if (err) ;
        else if (!item) ;
        else if (item.text) {
          if (temp == "orderNumber" && flag) {
            res["OrderNumber"] = item.text;
            temp = "";
            flag = false;
          }
          if (item.text == "Purchase Order Number" && !flag) {
            temp = "orderNumber";
            flag = true;
          }
          if (temp == "invoiceNumber" && flag) {
            res["InvoiceNumber"] = item.text;
            temp = "";
            flag = false;
          }
          if (item.text == "Invoice Number" && !flag) {
            temp = "invoiceNumber";
            flag = true;
          }
          if (flag && temp == "SHIP TO:") {
            res["Name"] = item.text;
            flag = false;
          }
          if (item.text == "SHIP TO:" && flag == false) {
            temp = "SHIP TO:";
            flag = true;
          }
          
          if (flag && temp == "address" && item.text != "Invoice Date") {
            temp = "address";
            k += item.text;
          }
          if (flag && temp == "Invoicedate") {
            temp = "";
            res["InvoiceDate"] = item.text;
            flag = false;
          }
          if (flag && temp == "address" && item.text == "Invoice Date") {
            res["Address"] = k;
            k = "";
            temp = "Invoicedate";
          }
          if (item.text == res["Name"] && flag == false) {
            temp = "address";
            flag = true;
          }
          if (temp == "order date" && flag) {
            res["OrderDate"] = item.text;
            temp = "";
            flag = false;
            // console.log(res);
          }
          if (item.text == "Order Date" && !flag) {
            temp = "order date";
            flag = true;
          }
          if (
            temp == "total1" &&
            flag &&
            item.text * 2 != Number(item.text) + Number(item.text)
          ) {
            k += item.text + " ";
          }
  
          if (
            temp == "total1" &&
            flag &&
            item.text * 2 == Number(item.text) + Number(item.text)
          ) {
            res["ProductTitle"] = k;
            res["HSN"] = item.text;
            temp = ""
            k = "";
          }
          if (temp == "total" && flag) {
            temp += String(item.text);
          }
          if (item.text == "Total" && !flag) {
            flag = true;
            temp = "total";
          }
          if (temp == res["TaxRate"] && flag) {
            res["TotalPrice"] = item.text;
            console.log(res);
            flag = false;
            temp = "";
          }
          if (item.text[0] == "I" && item.text[1] == "G" && item.text[2] == "S") {
            res["TaxRate"] = item.text;
            temp = item.text;
          }
        }
      });
    });
    console.log(" ");
        await fs.unlink('./invoice.pdf', function (err) {
            if (err) console.log("All Rows of the Sheets Traversed :-) ");
            r++;
            flow();
        });
        return;
}
  