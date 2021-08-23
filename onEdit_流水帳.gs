function onEdit(e) {  
  let ss = SpreadsheetApp.getActiveSheet();
  let cell = ss.getActiveCell();
  let col = cell.getColumn();
  if (ss.getName() == '流水帳' && col == 3) {
    let row = cell.getRow();
    let dtCell = ss.getRange(row,1)
    if (!cell.isBlank()) {
      if (dtCell.isBlank()) {
        dtCell.setValue(new Date()).setNumberFormat('yyyy/MM/dd');
      }
    } else {
      dtCell.setValue("")
    }
  }
};
