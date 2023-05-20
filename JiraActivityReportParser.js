const yargs  = require('yargs')
const fs  = require('fs')
const { XMLParser, XMLBuilder, XMLValidator} = require("fast-xml-parser");
const parser = new XMLParser();
const hideBin  = require('yargs/helpers').hideBin
const ObjectsToCsv = require('objects-to-csv');
const path = require('node:path'); 
const { convert } = require('html-to-text');
const reader = require('xlsx');

const options = {
  wordwrap: 1300,
  // ...
};

async function saveToCsv(entryArray) {
  let reportPath = path.basename(`.\\${argv.xmlPath}`, '.xml') + ".csv";
  return new ObjectsToCsv(entryArray).toDisk(reportPath, { allColumns: true });
}

async function saveToXls(entryArray) {
  let reportPath = path.basename(`.\\${argv.xmlPath}`, '.xml') + ".xls";  
  let workBook = reader.utils.book_new();
  const workSheet = reader.utils.json_to_sheet(entryArray);
  reader.utils.book_append_sheet(workBook, workSheet, reportPath);
  reader.writeFile(workBook, reportPath);
}

async function getEntrysFromXmlFile(fileName)
{
	const XMLdata = await fs.readFileSync(fileName, 'utf8');
	const data = parser.parse(XMLdata);
  let entryArray = []
  data.feed.entry.forEach(entry => {    
    try {
      entryArray.push({        
        published: entry['published'],
        updated: entry['updated'],
        authorName: entry['author']?.name??null,
        authorEmail: entry['author']?.email??null,
        activityTargetSummary: entry['activity:target']?.summary??null,
        activityObjectSummary: entry['activity:object']?.summary??null,
        content: convert(entry['content']??"", options)
        })        
    } catch (error) {
      console.error(error);
    }    
  });
  return entryArray;
}

const argv = yargs(hideBin(process.argv))
    .config('config', function (configPath) {
        return JSON.parse(fs.readFileSync(configPath, 'utf-8'))
    })
    .demandOption(['config'])   
    .argv;

// main.js / main.ts (the filename doesn't matter)
async function main() {
  let entryArray = await getEntrysFromXmlFile(argv.xmlPath)  
  //console.log(await new ObjectsToCsv(entryArray).toString());
  
  // saveToCsv(entryArray)
  await saveToXls(entryArray)  
}

if (require.main === module) {
  main();
}