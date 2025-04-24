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

var ExcelPath = "Template/Templete_Export_Memo.xlsx";
using var Reporter = new ExcelReport(ExcelPath);

var Buffer = File.ReadAllBytes("image.jpg");
Reporter
    .UsingSheet("專案備忘錄", SheetTrack =>
    {
        SheetTrack.UsingCell("A11").SetImage(Buffer);
    })
    .SaveAsXlsx("output");



