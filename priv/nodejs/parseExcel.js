var argv = require('argv');
var xlsx = require('node-xlsx').default;

var argv = argv.option({
  name: 'file',
  short: 'f',
  type: 'string',
  description: 'File Path',
  example: "'script --file=value' or 'script -f value'"
}).run();

var filePath = argv.options["file"];

if (filePath) {
  try {
    var worksheetsFromFile = xlsx.parse(filePath);
    console.log(`OK:EXCEL:${JSON.stringify(worksheetsFromFile)}`);
  } catch(error) {
    console.log(`ERROR:MESSAGE:${error.message}`);
  }
}
