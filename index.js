const excelToJson = require("convert-excel-to-json");
var fs = require("fs");
var json2xls = require("json2xls");

function converExceltoJson(inputFilePath, outputFilePath) {
  var result = new Promise((resolve, reject) => {
    const excelData = excelToJson({
      sourceFile: inputFilePath,
      header: {
        // Is the number of rows that will be skipped and will not be present at our result object. Counting from top to bottom
        rows: 1, // 2, 3, 4, etc.
      },
      columnToKey: {
        A: "id",
		B: "name",
		C: "age",
		D: "contact"
      },
    });
    const targetData = excelData[Object.keys(excelData)[0]]; //Select Sheet 1 only

    fs.writeFile(outputFilePath, JSON.stringify(targetData), function (err) {
      if (err) throw err;
      console.log("Saved!");
    });
    resolve(targetData);
  });
  return result;
}

async function converJsontoExcel(inputFilePath, outputFilePath) {
  var result = new Promise(async (resolve, reject) => {
    var jsonData;
    fs.readFile(inputFilePath, function (err, data) {
      if (err) {
        console.log("Err:", err);
      }

      jsonData = JSON.parse(data);
      resolve(jsonData);
    });
  });

  result.then((data) => {
    var xls = json2xls(data);
    fs.writeFileSync(outputFilePath, xls, "binary");
  });
}

converExceltoJson("./testData.xlsx", "./" + Date.now() + "_data.json");
converJsontoExcel("./data.json", "./" + Date.now() + "_converted.xlsx");
