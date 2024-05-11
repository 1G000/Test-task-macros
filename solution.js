let links = [];
let v;

// Копирует ссылки на видимые выделенные ячейки
(function CpyVis() {
  const activeSheet = Api.ActiveSheet;
  const selection = activeSheet.Selection;
  const sheetName = activeSheet.GetName();
  links = [];
  selection.ForEach(function (cell) {
    if (
      !activeSheet.GetRows(cell.GetRow()).GetHidden() &&
      !activeSheet.GetCols(cell.GetCol()).GetHidden()
    ) 
  });
  links.push(
    `'[${bookName}]${sheetName}'!${activeSheet
      .GetRange(${cell})
      .GetAddress(true, true, 'xlA1', false)}`,
  );
})();

// Вставляет ссылки на видимые выделенные ячейки
(function PstVis() {
  const activeSheet = Api.ActiveSheet;
  const selection = activeSheet.Selection;
  let i = 0;
  selection.ForEach(function (cell) {
    if (
      !activeSheet.GetRows(cell.GetRow()).GetHidden() &&
      !activeSheet.GetCols(cell.GetCol()).GetHidden()
    ) {
      i++;
      v = links[i - 1];
      cell.SetValue(=${v});
      v = '';
    }
  });
})();

// Прибавляет ссылки на видимые выделенные ячейки
(function PstVisPlus() {
  const activeSheet = Api.ActiveSheet;
  const selection = activeSheet.Selection;
  let j = 0;
  selection.ForEach(function (cell) {
    if (
      !activeSheet.GetRows(cell.GetRow()).GetHidden() &&
      !activeSheet.GetCols(cell.GetCol()).GetHidden()
    ) {
      j++;
      v = links[j - 1];
      cell.SetValue(${cell.GetValue()} + ${v});
      v = '';
    }
  });
})();

// На первом листе с А1 ячейки создает список листов со ссылками
(function getSheetsList() {
  const sheets = Api.GetSheets();
  const firstSheet = sheets[0];
  firstSheet.SetActive(true);
  sheets.forEach((sheet, index) => {
    let sheetName = sheet.GetName();
    firstSheet.SetHyperlink(
      A${index + 1},
      '',
      ${sheetName}!A1,
       ${sheetName},
    );
  });
})();

// В названии каждого листа заменяет пробелы на подчеркивания
(function renameSheets() {
  const sheets = Api.GetSheets();
  sheets.forEach((sheet) => {
    const sheetName = sheet.GetName();
    const newSheetName = sheetName.replaceAll(' ', '_');
    sheet.SetName(newSheetName);
  });
})();

// Отображает все скрытые листы в файле
(function makeHiddenSheetsVisible() {
  const hiddenSheets = Api.GetSheets().filter(
    (sheet) => sheet.GetVisible() === false,
  );
  hiddenSheets.forEach((sheet, index) => {
    sheet.SetVisible(true);
  });
})();
