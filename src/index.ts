import xlsx from "xlsx";
import fs from "fs";

const inputFilePath = __dirname + "/Февраль.xls";
const destFilePath = __dirname + "/test.csv";

const cellDataAccessKey = "Закупки";
const cellCostDataAccessKey = "__EMPTY";

const dateRegex = /\d\d\.\d\d\.\d\d\d\d/;

function main() {
  const xfile = xlsx.readFile(inputFilePath);
  const sheet = xfile.Sheets[xfile.SheetNames[0]];

  const json = xlsx.utils.sheet_to_json(sheet, {
    blankrows: false,
  });

  const csvWorkbook = xlsx.utils.book_new();

  const csvSheet = xlsx.utils.json_to_sheet(
    [
      {
        A: "Дата",
        B: "Контрагент",
        C: "Стоимость",
      },
    ],
    { skipHeader: true }
  );

  xlsx.utils.book_append_sheet(csvWorkbook, csvSheet);

  let lastDate = null;
  const jsonSheetData = [];
  const jsonWithoutUnknownRows = json.slice(7, -1);

  for (const entry of jsonWithoutUnknownRows) {

    if (dateRegex.test(entry[cellDataAccessKey])) {
      lastDate = entry[cellDataAccessKey];
      continue;
    }

    const company = entry[cellDataAccessKey];
    const cost = entry[cellCostDataAccessKey];

    jsonSheetData.push({
      A: lastDate,
      B: company,
      C: cost.toString().replace(".", ","),
    });
  }

  xlsx.utils.sheet_add_json(csvSheet, jsonSheetData, {
    origin: -1,
    skipHeader: true,
  });

  const csvOutput = xlsx.utils.sheet_to_csv(csvSheet, { FS: ";" });

  fs.writeFileSync(destFilePath, csvOutput, { encoding: "utf-8" });
}

main();
