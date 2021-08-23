function onEdit(e) {  
  let ss = SpreadsheetApp.getActiveSheet();
  if (ss.getName() == '檢定') {
    let cell = ss.getActiveCell();
    let col = cell.getColumn();
    let row = cell.getRow();
    if (col == 15 && row == 1) {
      ss.getRange('o1:o2').splitTextToColumns(',')
      ss.setActiveSelection('o1:as2')
    } else if (col == 46) {   
      let dtCell = ss.getRange(row - 1,col)
      if (!cell.isBlank() && dtCell.isBlank()) {
        dtCell.setValue(new Date()).setNumberFormat("M/d HH:mm");
      } else if (cell.isBlank() && !dtCell.isBlank()) {
        dtCell.setValue("")
      }
    }
  } else if (ss.getName() == '逐月已實現') {
    let cell = ss.getActiveCell();
     ss.getRange('a:a').splitTextToColumns(',')
  }
};
