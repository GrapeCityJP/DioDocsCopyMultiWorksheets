// See https://aka.ms/new-console-template for more information
Console.WriteLine("複数のワークシートをコピー");

// ワークブックを開く
var workbook = new GrapeCity.Documents.Excel.Workbook();
workbook.Open("multiworksheet.xlsx");

// Sheet1とSheet2を末尾にコピー
workbook.Worksheets[new string[] { "Sheet1", "Sheet2" }].Copy();

//// Sheet1とSheet2をSheet3の後にコピー
//workbook.Worksheets[new string[] { "Sheet1", "Sheet2" }].CopyAfter(workbook.Worksheets[2]);

//// Sheet1とSheet2をSheet1の前にコピー
//workbook.Worksheets[new string[] { "Sheet1", "Sheet2" }].CopyBefore(workbook.Worksheets[0]);

// Excelファイルとして保存
workbook.Save("CopyMultipleWorksheets.xlsx");