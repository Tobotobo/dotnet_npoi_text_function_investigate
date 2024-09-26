using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

static void exec(Action<XSSFWorkbook, ISheet, ICell, ICell, ICellStyle> action) {
    var workbook = new XSSFWorkbook();
    var sheet = workbook.CreateSheet("Sheet1");
    IRow row = sheet.CreateRow(0);
    ICell A1 = row.CreateCell(0);
    ICell B1 = row.CreateCell(1);

    // 日付のフォーマットを指定
    var creationHelper = workbook.GetCreationHelper();
    var cellStyle = workbook.CreateCellStyle();
    var dateFormat = creationHelper.CreateDataFormat().GetFormat("yyyy/mm/dd");
    cellStyle.DataFormat = dateFormat;

    // ケースの処理を実行
    action(workbook, sheet, A1, B1, cellStyle);

    // 数式を評価して結果を取得
    var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
    var value = evaluator.Evaluate(B1);

    // 結果を出力
    var s = "";
    switch(value.CellType) {
        case CellType.Blank:
            s = $"Blank";
            break;
        case CellType.Boolean:
            s = $"Boolean: {value.BooleanValue}";
            break;
        case CellType.Error:
            var errCode = value.ErrorValue;
            var errText = FormulaError.ForInt(errCode).String;
            s = $"Error: {errCode}: {errText}";
            break;
        case CellType.Formula:
            s = $"Formula: {value.StringValue}";
            break;
        case CellType.Numeric:
            s = $"Numeric: {value.NumberValue}";
            break;
        case CellType.String:
            s = $"String: {value.StringValue}";
            break;
        case CellType.Unknown:
            s = $"Unknown";
            break;
        default:
            s = "default";
            break;
    }
    Console.WriteLine(s);
}

Console.WriteLine("#1 日付型の値を直接渡した場合 → 変換され yyyy/mm/dd 形式で出力");
exec((workbook, sheet, A1, B1, cellStyle) => {
    A1.SetCellType(CellType.Numeric);
    A1.CellStyle = cellStyle;
    A1.SetCellValue("2012-12-31");

    B1.SetCellFormula("TEXT(NOW(),\"yyyy/mm/dd\")");
});

Console.WriteLine("#2 数値型の値を直接渡した場合 → 変換され yyyy/mm/dd 形式で出力");
exec((workbook, sheet, A1, B1, cellStyle) => {
    B1.SetCellFormula("TEXT(123,\"yyyy/mm/dd\")");
});

Console.WriteLine("#3 文字列型(日付)の値を直接渡した場合 → 変換されず元の値を出力");
exec((workbook, sheet, A1, B1, cellStyle) => {
    B1.SetCellFormula("TEXT(\"2012-12-31\",\"yyyy/mm/dd\")");
});

Console.WriteLine("#4 文字列型(数字)の値を直接渡した場合 → 変換されず元の値を出力");
exec((workbook, sheet, A1, B1, cellStyle) => {
    B1.SetCellFormula("TEXT(\"123\",\"yyyy/mm/dd\")");
});

Console.WriteLine("#5 文字列型(文字)の値を直接渡した場合 → 変換されず元の値を出力");
exec((workbook, sheet, A1, B1, cellStyle) => {
    B1.SetCellFormula("TEXT(\"abc\",\"yyyy/mm/dd\")");
});

Console.WriteLine("#6 日付型の値をセル経由で渡した場合 → 変換され yyyy/mm/dd 形式で出力");
exec((workbook, sheet, A1, B1, cellStyle) => {
    A1.SetCellType(CellType.Numeric);
    A1.CellStyle = cellStyle;
    // A1.SetCellValue("2012-12-31");
    A1.SetCellValue(DateTime.Now);

    B1.SetCellFormula("TEXT(A1,\"yyyy/mm/dd\")");
});

Console.WriteLine("#7 数値型の値をセル経由で渡した場合 → 変換され yyyy/mm/dd 形式で出力");
exec((workbook, sheet, A1, B1, cellStyle) => {
    A1.SetCellType(CellType.Numeric);
    A1.SetCellValue(123);

    B1.SetCellFormula("TEXT(A1,\"yyyy/mm/dd\")");
});

Console.WriteLine("#8 文字列型(日付)の値をセル経由で渡した場合(ハイフン) → #VALUE! エラーが発生");
exec((workbook, sheet, A1, B1, cellStyle) => {
    A1.SetCellType(CellType.String);
    A1.SetCellValue("2012-12-31");

    B1.SetCellFormula("TEXT(A1,\"yyyy/mm/dd\")");
});

Console.WriteLine("#9 文字列型(数字)の値をセル経由で渡した場合 → 変換され yyyy/mm/dd 形式で出力");
exec((workbook, sheet, A1, B1, cellStyle) => {
    A1.SetCellType(CellType.String);
    A1.SetCellValue("123");

    B1.SetCellFormula("TEXT(A1,\"yyyy/mm/dd\")");
});

Console.WriteLine("#10 文字列型(文字)の値をセル経由で渡した場合 → #VALUE! エラーが発生");
exec((workbook, sheet, A1, B1, cellStyle) => {
    A1.SetCellType(CellType.String);
    A1.SetCellValue("abc");

    B1.SetCellFormula("TEXT(A1,\"yyyy/mm/dd\")");
});

Console.WriteLine("#11 文字列型(日付)の値をセル経由で渡した場合(スラッシュ) → #VALUE! エラーが発生");
exec((workbook, sheet, A1, B1, cellStyle) => {
    A1.SetCellType(CellType.String);
    A1.SetCellValue("2012/12/31");

    B1.SetCellFormula("TEXT(A1,\"yyyy/mm/dd\")");
});

Console.WriteLine("#12 文字列型(日付)の値をセル経由で渡した場合(スラッシュ+日付書式) → #VALUE! エラーが発生");
exec((workbook, sheet, A1, B1, cellStyle) => {
    A1.SetCellType(CellType.Numeric);
    A1.CellStyle = cellStyle;
    A1.SetCellValue("2012/12/31");

    B1.SetCellFormula("TEXT(A1,\"yyyy/mm/dd\")");
});

Console.WriteLine("#13 文字列型(日付)の値をセル経由で渡した場合(ハイフン+日付書式) → #VALUE! エラーが発生");
exec((workbook, sheet, A1, B1, cellStyle) => {
    A1.SetCellType(CellType.Numeric);
    A1.CellStyle = cellStyle;
    A1.SetCellValue("2012-12-31");

    B1.SetCellFormula("TEXT(A1,\"yyyy/mm/dd\")");
});

Console.WriteLine("#14 文字列型(日付)の値をセル経由且つ日付に変換して渡した場合(ハイフン) → 変換され yyyy/mm/dd 形式で出力");
exec((workbook, sheet, A1, B1, cellStyle) => {
    A1.SetCellType(CellType.String);
    A1.SetCellValue("2012-12-31");

    B1.SetCellFormula("TEXT(DATEVALUE(A1),\"yyyy/mm/dd\")");
});

Console.WriteLine("#15 文字列型(日付)の値をセル経由且つ日付に変換して渡した場合(スラッシュ) → 変換され yyyy/mm/dd 形式で出力");
exec((workbook, sheet, A1, B1, cellStyle) => {
    A1.SetCellType(CellType.String);
    A1.SetCellValue("2012/12/31");

    B1.SetCellFormula("TEXT(DATEVALUE(A1),\"yyyy/mm/dd\")");
});

Console.WriteLine("#16 文字列型(日付)の値をセル経由且つ日付に変換して渡した場合(スラッシュ+日付書式) → 変換され yyyy/mm/dd 形式で出力");
exec((workbook, sheet, A1, B1, cellStyle) => {
    A1.SetCellType(CellType.Numeric);
    A1.CellStyle = cellStyle;
    A1.SetCellValue("2012/12/31");

    B1.SetCellFormula("TEXT(DATEVALUE(A1),\"yyyy/mm/dd\")");
});

Console.WriteLine("#17 文字列型(日付)の値をセル経由で渡した場合(ハイフン+日付書式) → 変換され yyyy/mm/dd 形式で出力");
exec((workbook, sheet, A1, B1, cellStyle) => {
    A1.SetCellType(CellType.Numeric);
    A1.CellStyle = cellStyle;
    A1.SetCellValue("2012-12-31");

    B1.SetCellFormula("TEXT(DATEVALUE(A1),\"yyyy/mm/dd\")");
});
