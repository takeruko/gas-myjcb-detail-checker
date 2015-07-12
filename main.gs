// MyJCB's URLs.
// You don't need to modify these constant values.
var LOGIN_URL = "https://my.jcb.co.jp/iss-pc/member/user_manage/Login";
var DETAIL_URL = "https://my.jcb.co.jp/iss-pc/member/details_inquiry/detail.html";
var PDF_DETAIL_URL = "https://my.jcb.co.jp/iss-pc/member/details_inquiry/detailDbPdf.html";

function GetLatestDetailFile() {
  // Get userId, password, mailTo from script properties.
  var properties = PropertiesService.getScriptProperties();
  var userId = properties.getProperty("userId");
  var password = properties.getProperty("password");
  var mailTo = properties.getProperty("mailTo");
  var detailFolderName = properties.getProperty("folderName");

  // Login
  var cookies = login(userId, password);
  
  // Before downloading detail pdf, must fetch the detail.html (I don't know why).
  getDetailBlob(cookies, DETAIL_URL, 1, "web");
  
  // Fetch and save the detail pdf.
  var outdir = getDetailFolder(detailFolderName, mailTo);
  var pdfDetailBlob = getDetailPdfBlob(cookies, 1);
  var pdfDetail = outdir.createFile(pdfDetailBlob);
  pdfDetail.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);

  // Fetch and save the detail csv.
  var csvDetailBlob = getDetailCsvBlob(cookies, 1);
  var csvDetail = outdir.createFile(csvDetailBlob);
  csvDetail.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);

  // Send report mail.  
  var yyyymm = pdfDetail.getName().slice(0, 4) + "年" + pdfDetail.getName().slice(4, 6) + "月"; // The detail pdf name probably <yyyymm>meisai.pdf .
  var mailBody = "JCBクレジットカード利用明細:" + yyyymm + "分を取得しました。\n";
      mailBody+= "下記URLからダウンロードしてください。\n\n";
      mailBody+= pdfDetail.getUrl();
  MailApp.sendEmail(mailTo, "JCBクレジットカード利用明細:" + yyyymm + "分" , mailBody);
}

function login(userid, password) {
  var options = {
    method : "post",
    followRedirects: false,
    contentType: "application/x-www-form-urlencoded",
    payload : {
      userId: userid,
      password: password
    }
  };

  var response = UrlFetchApp.fetch(LOGIN_URL, options);
  var headers = response.getAllHeaders();
  var cookies = [];
  if ( typeof headers['Set-Cookie'] !== 'undefined' ) {
    // Make sure that we are working with an array of cookies
    var cookies = typeof headers['Set-Cookie'] == 'string' ? [ headers['Set-Cookie'] ] : headers['Set-Cookie'];
    for (var i = 0; i < cookies.length; i++) {
      // We only need the cookie's value - it might have path, expiry time, etc here
      cookies[i] = cookies[i].split( ';' )[0];
    };
  }
  return cookies;    
}

function getDetailPdfBlob(cookies, detailMonth) { 

  var fileBlob = getDetailBlob(cookies, PDF_DETAIL_URL, detailMonth, "pdf");
  // If the detail file's extention is not ".pdf", set the default name to the detail file.
  if (fileBlob.getName().match(".pdf$") != ".pdf") {
    fileBlob.setName("meisai.pdf");
  }
  return fileBlob;
}

function getDetailCsvBlob(cookies, detailMonth) { 
  
  var fileBlob = getDetailBlob(cookies, DETAIL_URL, detailMonth, "csv");
  // If the detail file's extention is not ".pdf", set the default name to the detail file.
  if (fileBlob.getName().match(".csv$") != ".csv") {
    fileBlob.setName("meisai.csv");
  }
  return fileBlob;
}

function getDetailBlob(cookies, url, detailMonth, output) {
  
  var options = {
    method: "get",
    followRedirects: false,
    headers: {
      Cookie: cookies.join(';')
    }
  };
  
  var response = UrlFetchApp.fetch(url + "?detailMonth=" + detailMonth + "&output=" + output , options);
  var headers = response.getAllHeaders();  
  var fileBlob = response.getBlob();
  if ( typeof headers['Content-Disposition'] !== 'undefined' ) {
    // Get filename from Content-Disposition Header, and set the filename to the downloaded file.
    var filename = headers['Content-Disposition'].match('filename[^;=\n]*=["]*((["]).*?\2|[^;"\n]*)')[1];
    fileBlob.setName(filename);
  }
  return fileBlob;
}

function getDetailFolder(folderName, editorsAddress) {
  var folders = DriveApp.getFoldersByName(folderName);
  while (folders.hasNext()) {
    var folder = folders.next();
    if (folder.getName() == folderName) {
      return folder;
    }
  }
  
  // If folder is not found, create the folder.
  var newFolder = DriveApp.createFolder(folderName);
  if (typeof editorsAddress == "string") {
    // Add editors permission.
    newFolder.addEditors(editorsAddress.split(','));
  }
  return newFolder;
}
