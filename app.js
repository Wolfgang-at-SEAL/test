
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
  console.log ('Test name set to: "' + testname + '"');
}
var client = 'c:/seal/dpfclient.exe';
var file = 'c:/seal/test/testfiles/original.doc';
var testdir = 'c:/seal/test/testfiles/';
var targetfile = 'original';
var duration = 60 * 1000;
if (argv.duration){
  duration = argv.duration;
  console.log ('Test duration set to: ' + duration);
  duration = duration * 1000;
}
var nedb = require('nedb');
var db = new nedb({ filename: __dirname + '/testresult.db', autoload: true });
var start;
var date = new Date();
var starttime = Date.now();
var jobCount = 0;

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
   file = testdir + '/' + file;
   callDpfclient(file);
   
 });
});

function saveResult(data, callback) {
  //console.log('data: ' + JSON.stringify(data));
  db.insert(data, callback(data.file));
  
}

function callDpfclient (file) {
  console.time('client start');
  var args = ['-wf', 'dpf4convert.netconvertoffice', 
                //'-s', 'TARGETFILE=' + file,
		'-s', 'TARGETDIR=c:/seal/test/output',
		'-f', file
		];
  var now = Date.now();
  if (now - starttime > duration){
	  // end loop
	  return;
  }
  console.log ('starting job: ' + jobCount++ + '\n');
  var options = {};
  child = execFile(client,args,options,
  function (error, stdout, stderr) {
    console.log('stdout: ' + stdout);
    console.log('stderr: ' + stderr);
    if (error !== null) {
      console.log('exec error: ' + error);
    }
  var end = Date.now();         
  var elapsed = end - now;
  console.timeEnd('client start');
  saveResult({'file': file,
	      'start': now,
              'end': end,         
	      'elapsed': elapsed,
	      'testname': testname,
	      'jobcount' : jobCount,
  },callDpfclient);
  });
  }
