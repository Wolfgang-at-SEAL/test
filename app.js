var moment = require('moment');
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
	        '  --help      show this message\n' + 
		'  --name      name to be used in database for this test\n' +
		'  --duration  duration of the test in seconds (default: 60s)\n'
		);
return;
}

var testname = Date();
if (argv.name){
  testname = argv.name;
}
else {
  testname = moment().format("YYYY-MM-DD-HH-MM-SS"); 
  }
console.log ('Test name set to: "' + testname + '"');




var client = 'c:/seal/dpfclient.exe';
var file;
var testdir = path.join(__dirname, 'testfiles');
var database = path.join(__dirname, testname + '.db');
var excelfile = testname + '.xlsx';

var duration = 60 * 1000;
if (argv.duration){
  duration = argv.duration;
  console.log ('Test duration set to: ' + duration);
  duration = duration * 1000;
}

var nedb = require('nedb');
var db = new nedb({ filename: database, autoload: true });
var start;
var date = new Date();
var starttime = Date.now();
var jobCount = 0;
var clientsRunning = 0;

var results = [];
fs.readdir(testdir, function(err, files){
  if (err) {
    console.log('readdir returned error: ' + error);
    return;
  }
 var pending = files.length;
 if (!pending){
   // no files
   console.log('no files in testdir: ' + testdir);
   return; 
 }
 files.forEach(function(file) {
   file = testdir + path.sep + file;
   clientsRunning++;
   callDpfclient(file);
   
 });
});

function saveResult(data, callback) {
  //console.log('data: ' + JSON.stringify(data));
  db.insert(data, function (err, newData){
    var _callback = callback;	  
    if (err !== null) {
      console.log('nedb error: ' + err);
    }
    _callback(data.file);
  });
  
}

function callDpfclient (file) {
  console.time('client duration');
  var args = ['-wf', 'dpf4convert.netconvertoffice', 
                //'-s', 'TARGETFILE=' + file,
		'-s', 'TARGETDIR=' + __dirname + '/output',
		'-f', file
		];
  var now = Date.now();
  if (now - starttime > duration){
	  // end loop
	  eventEmitter.emit('startExport');
	  return;
  }
  var _jobCount = jobCount++;
  console.log ('starting job: ' + _jobCount + '\n');
  var options = {};
  child = execFile(client,args,options,
  function (error, stdout, stderr) {
    console.log('stdout: ' + stdout);
    //console.log('stderr: ' + stderr);
    if (error !== null) {
      console.log('exec error: ' + error);
    }
  var _start = now;  
  var end = Date.now();         
  var elapsed = end - _start;
  console.timeEnd('client duration');
  saveResult({
	  'testname': testname,
	  'start': _start,
          'end': end,         
	  'elapsed': elapsed,
	  'jobcount' : _jobCount,
	  'file': file,
  },callDpfclient);
  });
  }

function exportToExcel() {
  clientsRunning--;
  if (clientsRunning <= 0) {
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
        var workbook = excelbuilder.createWorkbook(__dirname, excelfile)
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
  } // if clientsRunning
} //exportToExcel

