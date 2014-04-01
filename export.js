
var excelbuilder = require('msexcel-builder');
var events = require('events');
var path = require('path');
var eventEmitter = new events.EventEmitter();

eventEmitter.on('startExport', exportToExcel);
var parseArgs = require('minimist');
var execFile = require('child_process').execFile
var child;

var fs = require('fs');

var argv = parseArgs(process.argv, {boolean: ['help']});

if (argv.help) {
console.log (process.argv[1] + '\n' +
	        '  --help      show this message\n' 
		);
return;
}

var excelfile = path.join(__dirname, 'dpftest.xlsx');
var nedb = require('nedb');
var db = new nedb({ filename: __dirname + '/testresult.db', autoload: true });
var jobCount = 0;

exportToExcel();

function exportToExcel() {
    console.log ('exporting database to file: ' + excelfile);
    // countrows
      // Find all documents in the collection
      db.find({}, function (err, docs) {
        var rows = docs.length;
        console.log ('exporting ' + rows + ' datasets');
        // detect columns from dataa
     
        var columns = ['testname', 'start', 'end', 'elapsed', 'file', 'jobcount'];
        var columnCount = columns.length;
        // create worksheet
        var workbook = excelbuilder.createWorkbook(__dirname, 'sample.xlsx')
        var sheet1 = workbook.createSheet('sheet1', columnCount, rows + 1);
        // write column titles
	// Sheet.set(col, row, str)
        for (var i = 1; i <= columnCount; i++) {
          sheet1.set(i, 1, columns[i-1]);
          };
        // export data
        var row = 1;
	var dataset;
	var value;
        for (var doc in docs) {
          row++;  
	  dataset = docs[doc];
          for (var i = 1; i <= columnCount; i++) {
            value = dataset[columns[i-1]];
  	    sheet1.set(i, row, value);
          }; //foreach column    
        }; // foreach doc
        // save file
        workbook.save(function(err){
          if (err) {
            console.log('error during export: ' + err);
            workbook.cancel();
          }
          else {
            console.log('export successfully completed');
          };
        }); //workbook.save
      }); //db.find
} //exportToExcel

