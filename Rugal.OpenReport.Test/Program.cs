//using SkiaSharp;

//using var PdfDom = SKDocument.CreatePdf("/output.pdf");
//using var PdfCanvas = PdfDom.BeginPage(595, 842);
//using var Paint = new SKPaint()
//{
//    Color = SKColors.Black,
//    TextSize = 20
//};

//PdfCanvas.DrawText("Hello, SkiaSharp PDF!", 50, 50, Paint);
//PdfDom.EndPage();

using Rugal.OpenReport.Services;

var ExcelPath = "Template/test.xlsx";
using var Reporter = new ExcelReport(ExcelPath);

Reporter
    .UsingSheet("Test", SheetTrack =>
    {
        SheetTrack.UsingRow(1, Row =>
        {
            Row.CopyTo(7);
        });

        SheetTrack.UsingRows(2, 3, Rows =>
        {
            Rows.CopyTo(9);
        });
    })
    .SaveAsXlsx("output");



