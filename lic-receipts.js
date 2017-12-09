let fs = require('fs'),
PDFParser = require("pdf2json");
_ = require("lodash");
var sanitize = require("sanitize-filename");
var moment = require('moment');
var Excel = require('exceljs');


var workbook = new Excel.Workbook();
var sheet = workbook.addWorksheet('My Sheet');

sheet.columns = [
     { header: "txnNo", key: "txnNo"},
	 { header: "policyNo", key: "policyNo"},
	 { header: "cusName", key: "cusName"},
	 { header: "agencyCode", key: "agencyCode"},
	 { header: "plan", key: "plan"},
	 { header: "term", key: "term"},
	 { header: "doc", key: "doc"},
	 { header: "instPrm", key: "instPrm"},
	 { header: "mode", key: "mode"},
	 { header: "sumAssured", key: "sumAssured"},
	 { header: "numInst", key: "numInst"},
	 { header: "dueFrom", key: "dueFrom"},
	 { header: "dueTo", key: "dueTo"},
	 { header: "totalPrm", key: "totalPrm"},
	 { header: "lateFee", key: "lateFee"},
	 { header: "cdFee", key: "cdFee"},
	 { header: "tax", key: "tax"},
	 { header: "totalAmt", key: "totalAmt"},
	 { header: "branch", key: "branch"},
	 { header: "nextDue", key: "nextDue"},
	 { header: "cashAmt", key: "cashAmt"},
	 { header: "chqAmt", key: "chqAmt"},
	 { header: "chqNo", key: "chqNo"},
	 { header: "chqDate", key: "chqDate"},
	 { header: "nextDue", key: "nextDue"},
	 { header: "merchantCode", key: "merchantCode"},
];


var fileNames = [];
var fileFolder = './lic/';
var fileCont = 0;

// let pdfParser = new PDFParser(this,1);



var loadPDF = function(filePath){
	console.log(filePath);
	console.log(fileNames.length);
	console.log(fileCont);
  if(fileNames.length === fileCont){
    //Insert in db and add any FINAL code, then return;
  }
  else{
	  
	let pdfParser = new PDFParser(this,1);
    pdfParser.loadPDF(filePath);

	var string = null;
	var astring = null;
	
	 
	pdfParser.on("pdfParser_dataError", errData => console.error(errData.parserError) );
	pdfParser.on("pdfParser_dataReady", pdfData => {
		
	// fs.writeFile("./aaa.txt", pdfParser.getRawTextContent());
	string = pdfParser.getRawTextContent();

	// console.log(string);
	// var string = "This\nstring\nhas\nmultiple\nlines.",
	
	var astring = string.split('\n');
	
	// txnNo, txnDate
	
	var txnNo = astring[2].replace(/\r?\n|\r/g,'').match(/([0-9]+)$/g)[0];
	
	var txnDate = moment(astring[3].replace(/\r?\n|\r/g,'').replace(/[^\d\s\/:]+/gi,''),["DD/MM/YYYY hh:mm:ss"]).format();
	
	console.log(txnDate);
	
	// 
	
	// policyNo, agencyCode extraction starts
				
	var match = /Next Due/;
	var foundon;
		
	_.forEach(astring, function (line, number) {
		// console.log("Line:"+number);
		// console.log(line);		
		if (match.exec(line))
			foundon = number;
	});

	// console.log(astring);
	   
	 var policyNo = astring[foundon+1].replace(/\r?\n|\r/g,'');
	 
	 var agencyCode = astring[foundon+2].replace(/\r?\n|\r/g,'').match(/([0-9]+)$/g)[0];
	 
	 var plan = astring[foundon+3].replace(/\r?\n|\r/g,'');
	 
	 var term = astring[foundon+4].replace(/\r?\n|\r/g,'');
	 
	 var doc = astring[foundon+5].replace(/\r?\n|\r/g,'');
	 
	 var instPrm = astring[foundon+6].replace(/\r?\n|\r/g,'');
	 
	 var mode = astring[foundon+7].replace(/\r?\n|\r/g,'');
	 
 	 var sumAssured = astring[foundon+8].replace(/\r?\n|\r/g,'');
	 
	 var numInst = astring[foundon+9].replace(/\r?\n|\r/g,'');
	 
	 var dueFrom = astring[foundon+10].replace(/\r?\n|\r/g,'');
	 
	 var dueTo = astring[foundon+11].replace(/\r?\n|\r/g,'');

 	 var totalPrm = astring[foundon+12].replace(/\r?\n|\r/g,'');
	 
 	 var lateFee = astring[foundon+13].replace(/\r?\n|\r/g,'');
	 
	 var cdFee = astring[foundon+14].replace(/\r?\n|\r/g,'');
	 
	 var tax = astring[foundon+15].replace(/\r?\n|\r/g,'');
	 
	 var totalAmt = astring[foundon+16].replace(/\r?\n|\r/g,'');
	 
	 var branch = astring[foundon+18].replace(/\r?\n|\r/g,'');
	 
	 var nextDue = astring[foundon+19	].replace(/\r?\n|\r/g,'');
	 
	 
	// policyNo, agencyCode extraction ends
	
	
	var match = /Cash/;
	var foundon;
		
	_.forEach(astring, function (line, number) {
		// console.log("Line:"+number);
		// console.log(line);		
		if (match.exec(line))
			foundon = number;
	});
	
	var cashAmt = astring[foundon].replace(/\r?\n|\r/g,'').replace(/[^\d\.]+/gi,'');
	
	var chqAmt = astring[foundon+1].replace(/\r?\n|\r/g,'').replace(/[^\d\.]+/gi,'');
	
	var chqNo = astring[foundon+2].replace(/\r?\n|\r/g,'').replace(/[^\d]+/gi,'');
	
	var chqDate = astring[foundon+3].replace(/\r?\n|\r/g,'').replace(/[^\d\/]+/gi,'');
	
	
	
	var match = /For Life Insurance Corporation of India/;
	var foundon;
		
	_.forEach(astring, function (line, number) {
		// console.log("Line:"+number);
		// console.log(line);		
		if (match.exec(line))
			foundon = number;
	});
	
	var merchantCode = astring[foundon+1].replace(/\r?\n|\r/g,'').replace(/[^\d]+/gi,'');
	
	 // cusName name extraction starts

	 var arr = /(policyholder )(.*)/.exec(string);

	 var cusName = arr[2].replace(/Shri\/Smt\.((Sri|Smt|Shri)\s)*/g,'');
	 
	 cusName = cusName.replace(/( A\/S;).*/i,"")
	 
	 // cusName name extraction ends
	 

	 console.log("=================");
	 
	console.log("txnNo: "+txnNo);
	console.log("policyNo: "+policyNo);
	console.log("cusName: "+cusName);
	console.log("agencyCode: "+agencyCode);
	console.log("plan: "+plan);
	console.log("term: "+term);
	console.log("doc: "+doc);
	console.log("instPrm: "+instPrm);
	console.log("mode: "+mode);
	console.log("sumAssured: "+sumAssured);
	console.log("numInst: "+numInst);
	console.log("dueFrom: "+dueFrom);
	console.log("dueTo: "+dueTo);
	console.log("totalPrm: "+totalPrm);
	console.log("lateFee: "+lateFee);
	console.log("cdFee: "+cdFee);
	console.log("tax: "+tax);
	console.log("totalAmt: "+totalAmt);
	console.log("branch: "+branch);
	console.log("nextDue: "+nextDue);
	console.log("cashAmt: "+cashAmt);
	console.log("chqAmt: "+chqAmt);
	console.log("chqNo: "+chqNo);
	console.log("chqDate: "+chqDate);
	console.log("merchantCode: "+merchantCode);
	 
	 
	 console.log("=================");
	 
	 
	 var rowValues = [];
 rowValues[1 ] = txnNo;
 rowValues[2 ] = policyNo;
 rowValues[3 ] = cusName;
 rowValues[4 ] = agencyCode;
 rowValues[5 ] = plan;
 rowValues[6 ] = term;
 rowValues[7 ] = doc;
 rowValues[8 ] = instPrm;
 rowValues[9 ] = mode;
 rowValues[10] = sumAssured;
 rowValues[11] = numInst;
 rowValues[12] = dueFrom;
 rowValues[13] = dueTo;
 rowValues[14] = totalPrm;
 rowValues[15] = lateFee;
 rowValues[16] = cdFee;
 rowValues[17] = tax;
 rowValues[18] = totalAmt;
 rowValues[19] = branch;
 rowValues[20] = nextDue;
 rowValues[21] = cashAmt;
 rowValues[22] = chqAmt;
 rowValues[23] = chqNo;
 rowValues[24] = chqDate;
 rowValues[25] = nextDue;
 rowValues[26] = merchantCode;
sheet.addRow(rowValues);


workbook.csv.writeFile("txns.csv")
    .then(function() {
        // done
    });
	 
	 
	 
	 
	 

	 var newFileName =  policyNo + '-' + cusName + '.pdf'
	 
	 var illegalRe = /[\/\?<>\\:\*\|":]/g;

	 newFileName = newFileName.replace(illegalRe,'');

	 newFileName = fileFolder + newFileName;
	 
	 console.log(newFileName);

	 
	fs.rename(filePath, newFileName, function(err) {
		if ( err ) console.log('ERROR: ' + err);
	});
	   
	   
	fileCont++; //increase the file counter
	loadPDF(fileFolder + fileNames[fileCont]); //parse the next file

	});
	  
	  
	  
	  
  }
};

 

// pdfParser.loadPDF("./aaa.pdf");

fs.readdir(fileFolder, function(err, files){
  for (var i = files.length - 1; i >= 0; i--) {
    if (files[i].indexOf('.pdf') !== -1){
      fileNames.push(files[i]);
    }
  }

  loadPDF(fileFolder + fileNames[fileCont]);
});




