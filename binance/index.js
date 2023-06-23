// XLSX is a global from the standalone script

async function analyze() {
  var output = [];

  const file = document.getElementById("inputFile").files[0];
  const data = await file.arrayBuffer();
  const input = XLSX.read(data);
  var sheet = input.SheetNames[0];
  output.push(GenerateRow("Date", "Description", "Expense", "Entrance"))
  var sheetLenght = input.Sheets[sheet]["!ref"].split(":")[1].substring(1);
  var fromDate = document.getElementById("inputFromDate").value?new Date(document.getElementById("inputFromDate").value):undefined
  if(fromDate){fromDate.setHours(0,0,0,0);}
  for (i = 2; i <= sheetLenght; i++) {
    var transactionDate = new Date(input.Sheets[sheet]["A" + i].v)
    if(!fromDate||fromDate<=transactionDate)
    output.push(GenerateRow(
      transactionDate.toLocaleDateString('it-IT') + " " + transactionDate.toLocaleTimeString('it-IT'),
      input.Sheets[sheet]["B" + i].v,
      input.Sheets[sheet]["C" + i].v,
      input.Sheets[sheet]["D" + i].v))
  }
  var wb = aoa_to_workbook(output);
  XLSX.writeFile(wb, "export.xlsx");
}

function GenerateRow(Date, Description, Expense, Entrance) {
  var row = [];
  row[0] = String(Date || "");
  row[1] = String(Description || "");
  row[2] = String(Expense || 0);
  row[3] = String(Entrance || 0);
  return row;
}

function aoa_to_workbook(data/*:Array<Array<any> >*/, opts)/*:Workbook*/ {
  return sheet_to_workbook(XLSX.utils.aoa_to_sheet(data, opts), opts);
}
function sheet_to_workbook(sheet/*:Worksheet*/, opts)/*:Workbook*/ {
  var n = opts && opts.sheet ? opts.sheet : "Sheet1";
  var sheets = {}; sheets[n] = sheet;
  return { SheetNames: [n], Sheets: sheets };
}
