function main(workbook: ExcelScript.Workbook) {
  //read data
  //prende il foglio attivo tabellato e crea un foglio nuovo SCHEMA
  let row_number:number = 0;
  const cols_number:number = 12;
  const table = workbook.getActiveWorksheet().getTables()[0];
  const texts = table.getRange().getTexts();

  let data_objects_arr: TableData[] = [];
  if (table.getRowCount() > 0) {
    data_objects_arr = returnObjectFromValues(texts);
  }
  row_number = data_objects_arr.length;
  //data_objects_arr è l'array di oggetti
  //write
  const newSheet = workbook.addWorksheet();
  newSheet.activate();
  //
  //newSheet.setName("SCHEMA");
  let max_columns:number = 0;
  let min_columns:number = 1000000;
  for(let i = 0; i<row_number; i++){ 
    let temp: number = +data_objects_arr[i]["X-base"];
    max_columns = (temp > max_columns) ? temp:max_columns;
    min_columns = (temp < min_columns) ? temp:min_columns;
  }
  let max_rows: number = 0;
  for (let i = 0; i<row_number; i++) {
    let temp: number = +data_objects_arr[i]["Y-altezza"];
    max_rows = (temp > max_rows) ? temp:max_rows;
  }
  
  let currentCell = workbook.getActiveCell();
  let activeCell = currentCell;
  let current_rows: number;
  let current_cols: number;
  for (let i = 0; i < max_rows; i++){
    for (let j = 0; j < max_columns; j++){
      current_rows= max_rows - (data_objects_arr[(i * max_columns) + j]["Y-altezza"] as unknown as number);
      current_cols = data_objects_arr[(i * max_columns) + j]["X-base"] as unknown as number - min_columns;
      currentCell = activeCell.getOffsetRange((current_rows)*4, (current_cols)*2);
      let targetRange = currentCell.getResizedRange(3,1);
      targetRange.setValues(buildCellObject(data_objects_arr[(i*max_columns)+j]));
      targetRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.diagonalDown).setStyle(ExcelScript.BorderLineStyle.none);
      targetRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
      targetRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setWeight(ExcelScript.BorderWeight.thin);
      targetRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setStyle(ExcelScript.BorderLineStyle.continuous);
      targetRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeTop).setWeight(ExcelScript.BorderWeight.thin);
      targetRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setStyle(ExcelScript.BorderLineStyle.continuous);
      targetRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeBottom).setWeight(ExcelScript.BorderWeight.thin);
      targetRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setStyle(ExcelScript.BorderLineStyle.continuous);
      targetRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeRight).setWeight(ExcelScript.BorderWeight.thin);
      targetRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideVertical).setStyle(ExcelScript.BorderLineStyle.none);
      targetRange.getFormat().getRangeBorder(ExcelScript.BorderIndex.insideHorizontal).setStyle(ExcelScript.BorderLineStyle.none);
    }
  }
//aux
  function returnObjectFromValues(values: string[][]): TableData[] {
    let objectArray: TableData[] = [];
    let objectKeys: string[] = [];
    for (let i = 0; i < values.length; i++) {
      if (i === 0) {
        objectKeys = values[i]
        continue;
      }
      if(values[i][0] == "")
        break;
      let object: { [key: string]: string } = {};
      for (let j = 0; j < cols_number; j++) {
        object[objectKeys[j]] = values[i][j]
      }

      objectArray.push(object as unknown as TableData);
    }
    return objectArray;
  }
  function buildCellObject(data:TableData): string[][]{
    return [["", data["REP"]], 
            [data["ID_plot"], ""],
            [data["VARIETY_NA"],""],
            [`Y: ${data["Y-altezza"]}`, `X: ${data["X-base"]}`]];
  }
  interface TableData {
    REP: string
    "BLOCK_NO._2": string
    TREATMENT_: string
    PLOT_NUMBE: string
    VARIETY_NA: string
    "LAST_YEAR'": string
    "LAST_YEAR'2": string
    ID_plot: string
    "Y-altezza": string
    "X-base": string
    blocco: string
  }
}
