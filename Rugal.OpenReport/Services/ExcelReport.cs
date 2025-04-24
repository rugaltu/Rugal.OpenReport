using ClosedXML.Excel;
using System.Text.RegularExpressions;

namespace Rugal.OpenReport.Services;
public class ExcelReport : IDisposable
{
    protected string TemplatePath { get; set; }
    protected string ExportPath { get; set; }
    public XLWorkbook Workbook { get; set; }
    public ExcelReport(string TemplatePath)
    {
        WithTemplate(TemplatePath);
    }
    public ExcelReport WithTemplate(string TemplatePath)
    {
        this.TemplatePath = TemplatePath;
        Workbook?.Dispose();
        Workbook = new XLWorkbook(this.TemplatePath);
        return this;
    }
    public ExcelReport WithExportPath(string ExportPath)
    {
        this.ExportPath = ExportPath;
        return this;
    }
    public ExcelReport UsingWorkbook(Action<IXLWorkbook> WorkbookFunc)
    {
        WorkbookFunc?.Invoke(Workbook);
        return this;
    }
    public SheetTrack ToSheet(string SheetName)
    {
        if (!Workbook.TryGetWorksheet(SheetName, out var WorkSheet))
            return null;

        var Sheet = new SheetTrack(this, WorkSheet);
        return Sheet;
    }
    public ExcelReport UsingSheet(string SheetName, Action<SheetTrack> UsingFunc = null)
    {
        var Track = ToSheet(SheetName);
        if (!Track.PrintRange.Any())
        {
            throw new Exception("PrintRange can not be empty, try after setup the PrintRange");
        }

        UsingFunc?.Invoke(Track);
        return this;
    }
    public ExcelReport SaveAsXlsx(string FileName, out string FullFileName)
    {
        FullFileName = VerifyFileName(FileName);
        if (ExportPath is not null)
            FullFileName = Path.Combine(ExportPath, FullFileName);

        Workbook.SaveAs(FullFileName);
        return this;
    }
    public ExcelReport SaveAsXlsx(string FileName)
    {
        SaveAsXlsx(FileName, out _);
        return this;
    }
    public ExcelReport SaveAsXlsx()
    {
        Workbook.SaveAs(VerifyFileName(ExportPath));
        return this;
    }
    public void Dispose()
    {
        Workbook?.Dispose();
        GC.SuppressFinalize(this);
    }
    protected virtual string VerifyFileName(string FileName, string Extention = "xlsx")
    {
        if (FileName is null)
            throw new Exception("[VerifyFileName]: Output file name cannot be null");
        if (!Regex.IsMatch(FileName, @$"(?i)\.{Extention}$", RegexOptions.IgnoreCase))
            FileName += $".{Extention}";
        return FileName;
    }
}