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
var ImagePath = "image4";
using var Reporter = new ExcelReport(ExcelPath);

var Model = new TestModel()
{
    ProjectName = "測試ProjectNameAAA",
    Groups = new List<TestGroup>()
    {
        new TestGroup()
        {
            GroupName = "群組A",
            Description = "AAA",
            Images = new List<ImageItem>()
            {
                new ImageItem()
                {
                    Name = "ImageA-1",
                },
                new ImageItem()
                {
                    Name = "ImageA-2",
                },
            }
        },
        new TestGroup()
        {
            GroupName = "群組B",
            Description = "BBB",
            Images = new List<ImageItem>()
            {
                new ImageItem()
                {
                    Name = "ImageB-1",
                },
                new ImageItem()
                {
                    Name = "ImageB-2",
                },
            }
        },
    }
};
Reporter
    .UsingSheet("專案備忘錄", SheetTrack =>
    {
        SheetTrack.UsingCell("A11", Cell =>
        {
            var Buffer = File.ReadAllBytes(ImagePath);
            Cell.SetImage(Buffer);
        });
        //SheetTrack.UsingRange("A9", "D11")
        //    .CopyTo("A15");

        //SheetTrack.WithStore(Model);
        //SheetTrack.WriteStoreBinding();
    })
    .SaveAsXlsx("output");





class ImageItem
{
    public string Name { get; set; }
}
class TestGroup
{
    public string GroupName { get; set; }
    public string Description { get; set; }
    public List<ImageItem> Images { get; set; }
}

class TestModel
{
    public string ProjectName { get; set; }
    public List<TestGroup> Groups { get; set; }
}

