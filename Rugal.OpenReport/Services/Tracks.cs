using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using ImageMagick;
using Rugal.OpenReport.Models;
using Rugal.OpenReport.StoreBinding;

namespace Rugal.OpenReport.Services;
public abstract class TrackBase
{
    public IXLWorksheet Worksheet { get; protected set; }
    public IXLPrintAreas PrintRange => Worksheet.PageSetup.PrintAreas;
    public ExcelReport Reporter { get; protected set; }
    public TrackBase(ExcelReport Reporter, IXLWorksheet Worksheet)
    {
        this.Reporter = Reporter;
        this.Worksheet = Worksheet;
    }
    public TrackBase(TrackBase RootTrack) : this(RootTrack.Reporter, RootTrack.Worksheet)
    {
    }
    protected virtual RowTrack ToRow(int Row)
    {
        var GetRow = Worksheet.Row(Row);
        return new RowTrack(this, GetRow);
    }
    protected virtual RowsRangeTrack ToRows(int StartRow, int EndRow)
    {
        var Rows = Worksheet.Rows(StartRow, EndRow);
        var Track = new RowsRangeTrack(this, Rows);
        return Track;
    }
    protected virtual CellTrack ToCell(int Row, int Column)
    {
        var Cell = Worksheet.Cell(Row, Column);
        return new CellTrack(this, Cell);
    }
    protected virtual CellTrack ToCell(int Row, string ColumnAddress)
    {
        var Cell = Worksheet.Cell(Row, ColumnAddress);
        return new CellTrack(this, Cell);
    }
    protected virtual CellTrack ToCell(string Address)
    {
        var Cell = Worksheet.Cell(Address);
        return new CellTrack(this, Cell);
    }
    protected virtual RangeTrack ToRange(CellTrack StartTrack, CellTrack EndTrack)
    {
        var GetRange = Worksheet.Range(StartTrack.Address, EndTrack.Address);
        var Track = new RangeTrack(this, GetRange);
        return Track;
    }
    protected virtual RangeTrack ToRange(string StartAddress, string EndAddress)
    {
        var GetRange = Worksheet.Range(StartAddress, EndAddress);
        var Track = new RangeTrack(this, GetRange);
        return Track;
    }
    protected virtual RangeTrack ToRange(int StartRow, int StartColumn, int EndRow, int EndColumn)
    {
        var GetRange = Worksheet.Range(StartRow, StartColumn, EndRow, EndColumn);
        var Track = new RangeTrack(this, GetRange);
        return Track;
    }
    public virtual TrackBase CopyMerge(CopyMergeOption Option)
    {
        //find any merge range from (StartRow, StartColumn) to (EndRow, EndColumn)
        //and paste to (TargetRow, TargetColumn)
        var MergeRanges = Worksheet.MergedRanges.ToArray();
        foreach (var MergeRange in MergeRanges)
        {
            var MergeStartRow = MergeRange.FirstRow().RowNumber();
            var MergeStartColumn = MergeRange.FirstColumn().ColumnNumber();

            var MergeEndRow = MergeRange.LastRow().RowNumber();
            var MergeEndColumn = MergeRange.LastColumn().ColumnNumber();

            if (Option.StartRow != 0 && MergeStartRow < Option.StartRow)
                continue;

            if (Option.StartColumn != 0 && MergeStartColumn < Option.StartColumn)
                continue;

            if (Option.EndRow != 0 && MergeEndRow > Option.EndRow)
                continue;

            if (Option.EndColumn != 0 && MergeEndColumn > Option.EndColumn)
                continue;

            var MoveMergeStartRow = MergeStartRow - Option.StartRow;
            var MoveMergeStartColumn = MergeStartColumn - Option.StartColumn;
            var MoveMergeEndRow = MergeEndRow - Option.StartRow;
            var MoveMergeEndColumn = MergeEndColumn - Option.StartColumn;

            var SetMergeStartRow = Option.TargetRow + MoveMergeStartRow;
            var SetMergeStartColumn = Option.TargetColumn + MoveMergeStartColumn;
            var SetMergeEndRow = Option.TargetRow + MoveMergeEndRow;
            int SetMergeEndColumn = Option.TargetColumn + MoveMergeEndColumn;

            Worksheet
                .Range(SetMergeStartRow, SetMergeStartColumn, SetMergeEndRow, SetMergeEndColumn)
                .Merge();
        }
        return this;
    }
    public virtual double ConvertWidthToPixels(double Width)
    {
        if (Width == 0)
            return 0;

        if (Width <= 1)
            return Width * 12;

        return (Width - 1) * 7 + 5;
    }
    public virtual double ConvertHeightToPixels(double Height)
    {
        return Height * 96 / 72;
    }
    public virtual double ConvertPixelsToWidth(double Pixels)
    {
        if (Pixels <= 5)
            return 1.0;

        return (Pixels - 5) / 7.0 + 1;
    }
    public virtual double ConvertPixelsToHeight(double Pixels)
    {
        return Pixels * 72.0 / 96.0;
    }

    protected virtual IXLPicture AddImage(Stream Stream)
    {
        var SupportFormats = new[]
        {
            MagickFormat.Jpeg,
            MagickFormat.Jpg,
            MagickFormat.Png,
            MagickFormat.Gif,
            MagickFormat.Bmp,
            MagickFormat.Tiff,
            MagickFormat.Tif,
        };
        var ConvertImage = new MagickImage(Stream);
        if (SupportFormats.Contains(ConvertImage.Format))
        {
            var SupportImage = Worksheet.AddPicture(Stream);
            Stream.Dispose();
            return SupportImage;
        }

        var ConvertStream = new MemoryStream();
        ConvertImage.Write(ConvertStream, MagickFormat.Png);
        var JpegImage = Worksheet.AddPicture(ConvertStream);

        ConvertStream?.Dispose();
        ConvertImage?.Dispose();
        return JpegImage;
    }
}
public class SheetTrack : TrackBase
{
    public BindingSet BindingSet { get; set; }
    public object Store { get; set; }
    public SheetTrack(ExcelReport Reporter, IXLWorksheet Worksheet) : base(Reporter, Worksheet)
    {
        BindingSet = new BindingSet(this);
    }
    public SheetTrack WithStore<TStore>(TStore Store)
    {
        this.Store = Store;
        return this;
    }
    public SheetTrack WriteStoreBinding()
    {
        BindingSet.WriteBinding();
        return this;
    }
    public RowTrack UsingRow(int Row, Action<RowTrack> UsingFunc = null)
    {
        var Track = ToRow(Row);
        UsingFunc?.Invoke(Track);
        return Track;
    }
    public RowsRangeTrack UsingRowsRange(int StartRow, int EndRow, Action<RowsRangeTrack> UsingFunc = null)
    {
        var Track = ToRows(StartRow, EndRow);
        UsingFunc?.Invoke(Track);
        return Track;
    }
    public CellTrack UsingCell(int Row, int Column, Action<CellTrack> UsingFunc = null)
    {
        var Track = ToCell(Row, Column);
        UsingFunc?.Invoke(Track);
        return Track;
    }
    public CellTrack UsingCell(int Row, string ColumnAddress, Action<CellTrack> UsingFunc = null)
    {
        var Cell = ToCell(Row, ColumnAddress);
        UsingFunc?.Invoke(Cell);
        return Cell;
    }
    public CellTrack UsingCell(string Address, Action<CellTrack> UsingFunc = null)
    {
        var Cell = ToCell(Address);
        UsingFunc?.Invoke(Cell);
        return Cell;
    }
    public RangeTrack UsingRange(CellTrack StartTrack, CellTrack EndTrack, Action<RangeTrack> UsingFunc = null)
    {
        var Track = ToRange(StartTrack, EndTrack);
        UsingFunc?.Invoke(Track);
        return Track;
    }
    public RangeTrack UsingRange(string StartAddress, string EndAddress, Action<RangeTrack> UsingFunc = null)
    {
        var Track = ToRange(StartAddress, EndAddress);
        UsingFunc?.Invoke(Track);
        return Track;
    }
    public RangeTrack UsingRange(int StartRow, int StartColumn, int EndRow, int EndColumn, Action<RangeTrack> UsingFunc = null)
    {
        var Track = ToRange(StartRow, StartColumn, EndRow, EndColumn);
        UsingFunc?.Invoke(Track);
        return Track;
    }
}
public class RowTrack : TrackBase
{
    public IXLRow Row { get; protected set; }
    public int RowNumber => Row.RowNumber();
    public RowTrack(TrackBase RootTrack, IXLRow Row) : base(RootTrack)
    {
        this.Row = Row;
    }
    public RowTrack MoveRow(int RelativeRow, Action<RowTrack> MoveFunc = null)
    {
        var Track = ToRow(RowNumber + RelativeRow);
        MoveFunc?.Invoke(Track);
        return Track;
    }
    public CellTrack UsingColumn(int Column, Action<CellTrack> UsingFunc = null)
    {
        var GetCell = Row.Cell(Column);
        var Track = new CellTrack(this, GetCell);
        UsingFunc?.Invoke(Track);
        return Track;
    }
    public CellTrack UsingColumn(string ColumnAddress, Action<CellTrack> UsingFunc = null)
    {
        var GetCell = Row.Cell(ColumnAddress);
        var Track = new CellTrack(this, GetCell);
        UsingFunc?.Invoke(Track);
        return Track;
    }
    public RowTrack UsingCopyTo(CopyRowOption Option, Action<RowTrack> UsingFunc = null)
    {
        var TargetRowNumber = 0;
        if (Option.PositionType == PositionTypes.Assign)
            TargetRowNumber = Option.TargetRow;
        else if (Option.PositionType == PositionTypes.Move)
            TargetRowNumber = RowNumber + Option.TargetRow;

        var TargetRow = Worksheet.Row(TargetRowNumber);
        if (Option.PasteType != PasteTypes.Overwrite)
        {
            if (Option.PasteType == PasteTypes.InsertBefore)
                TargetRow = TargetRow.InsertRowsAbove(1).First();
            else if (Option.PasteType == PasteTypes.InsertAfter)
                TargetRow = TargetRow.InsertRowsBelow(1).First();

            TargetRowNumber = TargetRow.RowNumber();
        }

        TargetRow.Style = Row.Style;
        TargetRow.Height = Row.Height;

        var StartUsedColumn = PrintRange.First().FirstColumn().ColumnNumber();
        var EndUsedColumn = PrintRange.First().LastColumn().ColumnNumber();

        for (var i = StartUsedColumn; i <= EndUsedColumn; i++)
        {
            var SourceCell = Row.Cell(i);
            var TargetCell = TargetRow.Cell(i);
            TargetCell.Value = SourceCell.Value;
            TargetCell.Style = SourceCell.Style;
        }

        CopyMerge(new CopyMergeOption()
        {
            StartRow = RowNumber,
            EndRow = RowNumber,
            TargetRow = TargetRowNumber,
        });

        var TargetRowTrack = new RowTrack(this, TargetRow);
        UsingFunc?.Invoke(TargetRowTrack);
        return TargetRowTrack;
    }
    public RowTrack UsingCopyTo(int TargetRow, Action<CopyRowOption> OptionFunc = null, Action<RowTrack> UsingFunc = null)
    {
        var Option = new CopyRowOption()
        {
            TargetRow = TargetRow,
        };
        OptionFunc?.Invoke(Option);
        var Track = UsingCopyTo(Option, UsingFunc);
        return Track;
    }
    public RowTrack UsingCopyTo(RowTrack SourceTrack, PasteTypes PasteType = PasteTypes.Overwrite, Action<RowTrack> UsingFunc = null)
    {
        var Track = UsingCopyTo(new CopyRowOption()
        {
            PositionType = PositionTypes.Assign,
            TargetRow = SourceTrack.RowNumber,
            PasteType = PasteType,
        }, UsingFunc);
        return Track;
    }
    public RowTrack UsingCopyTo(RowsRangeTrack SourceTrack, PasteTypes PasteType = PasteTypes.Overwrite, Action<RowTrack> UsingFunc = null)
    {
        var TargetRow = PasteType switch
        {
            PasteTypes.Overwrite => SourceTrack.StartRow,
            PasteTypes.InsertBefore => SourceTrack.StartRow,
            PasteTypes.InsertAfter => SourceTrack.EndRow,
            _ => SourceTrack.StartRow,
        };
        var Track = UsingCopyTo(new CopyRowOption()
        {
            PositionType = PositionTypes.Assign,
            TargetRow = TargetRow,
            PasteType = PasteType,
        }, UsingFunc);
        return Track;
    }
    public RowTrack UsingCopyTo(CellTrack SourceTrack, PasteTypes PasteType = PasteTypes.Overwrite, Action<RowTrack> UsingFunc = null)
    {
        var Track = UsingCopyTo(new CopyRowOption()
        {
            PositionType = PositionTypes.Assign,
            TargetRow = SourceTrack.RowNumber,
            PasteType = PasteType,
        }, UsingFunc);
        return Track;
    }
    public RowTrack CopyTo(CopyRowOption Option)
    {
        UsingCopyTo(Option);
        return this;
    }
    public RowTrack CopyTo(int TargetRow, Action<CopyRowOption> OptionFunc = null)
    {
        UsingCopyTo(TargetRow, OptionFunc);
        return this;
    }
    public RowTrack CopyTo(RowTrack SourceTrack, PasteTypes PasteType = PasteTypes.Overwrite)
    {
        UsingCopyTo(SourceTrack, PasteType);
        return this;
    }
    public RowTrack CopyTo(RowsRangeTrack SourceTrack, PasteTypes PasteType = PasteTypes.Overwrite)
    {
        UsingCopyTo(SourceTrack, PasteType);
        return this;
    }
    public RowTrack CopyTo(CellTrack SourceTrack, PasteTypes PasteType = PasteTypes.Overwrite)
    {
        UsingCopyTo(SourceTrack, PasteType);
        return this;
    }
    public void Delete()
    {
        Row.Delete();
    }
}
public class RowsRangeTrack : TrackBase
{
    public IXLRows Rows { get; protected set; }
    public IXLRow[] RowList => [.. Rows];
    public int RowsCount => Rows.Count();
    public int StartRow => Rows.First().RowNumber();
    public int EndRow => Rows.Last().RowNumber();
    public RowsRangeTrack(TrackBase RootTrack, IXLRows Rows) : base(RootTrack)
    {
        this.Rows = Rows;
    }
    public RowTrack UsingRow(int Row, Action<RowTrack> UsingFunc = null)
    {
        var RowsCount = Rows.Count();
        if (Row > RowsCount)
            throw new Exception($"[UsingRow]: Row number {Row} cannot be greater than rows count {RowsCount}");

        var GetRow = Rows.ToArray()[Row - 1];
        var Track = new RowTrack(this, GetRow);
        UsingFunc?.Invoke(Track);
        return Track;
    }
    public RowsRangeTrack UsingCopyTo(CopyRowOption Option, Action<RowsRangeTrack> UsingFunc = null)
    {
        var TargetRowNumber = 0;
        if (Option.PositionType == PositionTypes.Assign)
            TargetRowNumber = Option.TargetRow;
        else if (Option.PositionType == PositionTypes.Move)
        {
            if (Option.TargetRow < 0)
                TargetRowNumber = StartRow + Option.TargetRow;
            else
                TargetRowNumber = EndRow + Option.TargetRow;
        }

        var TargetRow = Worksheet.Row(TargetRowNumber);
        var InsertedRows = Worksheet.Rows(TargetRowNumber, TargetRowNumber + RowsCount - 1);
        if (Option.PasteType != PasteTypes.Overwrite)
        {
            var InsertRowCount = RowsCount;
            if (Option.PasteType == PasteTypes.InsertBefore)
                InsertedRows = TargetRow.InsertRowsAbove(InsertRowCount);
            else if (Option.PasteType == PasteTypes.InsertAfter)
                InsertedRows = TargetRow.InsertRowsBelow(InsertRowCount);

            TargetRow = InsertedRows.First();
            TargetRowNumber = TargetRow.RowNumber();
        }

        foreach (var Row in Rows)
        {
            new RowTrack(this, Row).UsingCopyTo(new CopyRowOption()
            {
                TargetRow = TargetRow.RowNumber(),
                PasteType = PasteTypes.Overwrite,
                PositionType = PositionTypes.Assign,
            });
            TargetRow = TargetRow.RowBelow();
        }

        CopyMerge(new CopyMergeOption()
        {
            StartRow = StartRow,
            EndRow = EndRow,
            TargetRow = TargetRowNumber,
        });

        var InsertedRowsRangeTrack = new RowsRangeTrack(this, InsertedRows);
        UsingFunc?.Invoke(InsertedRowsRangeTrack);
        return InsertedRowsRangeTrack;
    }
    public RowsRangeTrack UsingCopyTo(int TargetRow, Action<CopyRowOption> OptionFunc = null, Action<RowsRangeTrack> UsingFunc = null)
    {
        var Option = new CopyRowOption()
        {
            TargetRow = TargetRow,
        };
        OptionFunc?.Invoke(Option);
        var Track = UsingCopyTo(Option, UsingFunc);
        return Track;
    }
    public RowsRangeTrack UsingCopyTo(RowTrack SourceTrack, PasteTypes PasteType = PasteTypes.Overwrite, Action<RowsRangeTrack> UsingFunc = null)
    {
        var Track = UsingCopyTo(new CopyRowOption()
        {
            PositionType = PositionTypes.Assign,
            TargetRow = SourceTrack.RowNumber,
            PasteType = PasteType,
        }, UsingFunc);
        return Track;
    }
    public RowsRangeTrack UsingCopyTo(RowsRangeTrack SourceTrack, PasteTypes PasteType = PasteTypes.Overwrite, Action<RowsRangeTrack> UsingFunc = null)
    {
        var TargetRow = PasteType switch
        {
            PasteTypes.Overwrite => SourceTrack.StartRow,
            PasteTypes.InsertBefore => SourceTrack.StartRow,
            PasteTypes.InsertAfter => SourceTrack.EndRow,
            _ => SourceTrack.StartRow,
        };
        var Track = UsingCopyTo(new CopyRowOption()
        {
            PositionType = PositionTypes.Assign,
            TargetRow = TargetRow,
            PasteType = PasteType,
        }, UsingFunc);
        return Track;
    }
    public RowsRangeTrack UsingCopyTo(CellTrack SourceTrack, PasteTypes PasteType = PasteTypes.Overwrite, Action<RowsRangeTrack> UsingFunc = null)
    {
        var Track = UsingCopyTo(new CopyRowOption()
        {
            PositionType = PositionTypes.Assign,
            TargetRow = SourceTrack.RowNumber,
            PasteType = PasteType,
        }, UsingFunc);
        return Track;
    }
    public RowsRangeTrack CopyTo(CopyRowOption Option)
    {
        UsingCopyTo(Option);
        return this;
    }
    public RowsRangeTrack CopyTo(int TargetRow, Action<CopyRowOption> OptionFunc = null)
    {
        UsingCopyTo(TargetRow, OptionFunc);
        return this;
    }
    public RowsRangeTrack CopyTo(RowTrack SourceTrack, PasteTypes PasteType = PasteTypes.Overwrite)
    {
        UsingCopyTo(SourceTrack, PasteType);
        return this;
    }
    public RowsRangeTrack CopyTo(RowsRangeTrack SourceTrack, PasteTypes PasteType = PasteTypes.Overwrite)
    {
        UsingCopyTo(SourceTrack, PasteType);
        return this;
    }
    public RowsRangeTrack CopyTo(CellTrack SourceTrack, PasteTypes PasteType = PasteTypes.Overwrite)
    {
        UsingCopyTo(SourceTrack, PasteType);
        return this;
    }
    public void Delete()
    {
        Rows.Delete();
    }
}
public class RangeTrack : TrackBase
{
    public IXLRange Range { get; protected set; }
    public IXLAddress StartAddress => Range.RangeAddress.FirstAddress;
    public IXLAddress EndAddress => Range.RangeAddress.LastAddress;
    public RangeTrack(TrackBase RootTrack, IXLRange Range) : base(RootTrack)
    {
        this.Range = Range;
    }
    public RangeTrack MoveCell(int RelativeRow, int RelativeColumn, Action<CellTrack> MoveFunc = null)
    {
        var TargetRow = StartAddress.RowNumber + RelativeRow;
        var TargetColumn = StartAddress.ColumnNumber + RelativeColumn;
        var Track = ToCell(TargetRow, TargetColumn);
        MoveFunc?.Invoke(Track);
        return this;
    }
    public RangeTrack MoveCellFromEnd(int RelativeRow, int RelativeColumn, Action<CellTrack> MoveFunc = null)
    {
        var TargetRow = EndAddress.RowNumber + RelativeRow;
        var TargetColumn = EndAddress.ColumnNumber + RelativeColumn;
        var Track = ToCell(TargetRow, TargetColumn);
        MoveFunc?.Invoke(Track);
        return this;
    }
    public RangeTrack CopyTo(IXLCell PositionCell)
    {
        var TargetRow = PositionCell.WorksheetRow();
        var TargetColumn = PositionCell.WorksheetColumn();

        var StartColumn = StartAddress.ColumnNumber;
        var EndColumn = EndAddress.ColumnNumber;

        foreach (var Row in Range.Rows())
        {
            TargetRow.Style = Row.Style;
            TargetRow.Height = Row.WorksheetRow().Height;
            for (var i = StartColumn; i <= EndColumn; i++)
            {
                var SourceCell = Row.Cell(i);
                var TargetCell = TargetRow.Cell(i);
                TargetCell.Value = SourceCell.Value;
                TargetCell.Style = SourceCell.Style;
            }

            TargetRow = TargetRow.RowBelow();
        }

        CopyMerge(new CopyMergeOption()
        {
            StartRow = StartAddress.RowNumber,
            StartColumn = StartAddress.ColumnNumber,
            EndRow = EndAddress.RowNumber,
            EndColumn = EndAddress.ColumnNumber,
            TargetRow = TargetRow.RowNumber(),
            TargetColumn = TargetColumn.ColumnNumber(),
        });
        return this;
    }
    public RangeTrack CopyTo(CellTrack CellTrack)
    {
        CopyTo(CellTrack.Cell);
        return this;
    }
    public RangeTrack CopyTo(string Address)
    {
        CopyTo(Worksheet.Cell(Address));
        return this;
    }
    public RangeTrack CopyTo(int Row, int Column)
    {
        CopyTo(Worksheet.Cell(Row, Column));
        return this;
    }
}
public class CellTrack : TrackBase
{
    public IXLCell Cell { get; protected set; }
    public IXLAddress Address => Cell.Address;
    public int RowNumber => Cell.Address.RowNumber;
    public int ColumnNumber => Cell.Address.ColumnNumber;
    public CellTrack(TrackBase RootTrack, IXLCell Cell) : base(RootTrack)
    {
        this.Cell = Cell;
    }
    public RowTrack UsingRow(Action<RowTrack> UsingFunc = null)
    {
        var Track = ToRow(Address.RowNumber);
        UsingFunc?.Invoke(Track);
        return Track;
    }
    public CellTrack Move(int RelativeRow, int RelativeColumn, Action<CellTrack> MoveFunc = null)
    {
        var Track = ToCell(RowNumber + RelativeRow, ColumnNumber + RelativeColumn);
        MoveFunc?.Invoke(Track);
        return Track;
    }
    public CellTrack MoveRow(int RelativeRow, Action<CellTrack> MoveFunc = null)
    {
        var Track = Move(RelativeRow, 0, MoveFunc);
        return Track;
    }
    public CellTrack MoveColumn(int RelativeColumn, Action<CellTrack> MoveFunc = null)
    {
        var Track = Move(0, RelativeColumn, MoveFunc);
        return Track;
    }
    public CellTrack SetImage(Stream Stream)
    {
        var Image = AddImage(Stream);
        var CellWidthPx = (int)GetCellWidthPixels();
        var CellHeightPx = (int)GetCellHeightPixels();

        if (Cell.IsMerged())
        {
            var MergedRange = Cell.MergedRange();
            CellWidthPx = (int)MergedRange.Columns()
                .Sum(Item =>
                {
                    var Width = Worksheet.Column(Item.ColumnNumber()).Width;
                    var WidthPx = ConvertWidthToPixels(Width);
                    return WidthPx;
                });

            CellHeightPx = (int)MergedRange.Rows()
                .Sum(Item =>
                {
                    var Height = Worksheet.Row(Item.RowNumber()).Height;
                    var HeightPx = ConvertHeightToPixels(Height);
                    return HeightPx;
                });

            MergedRange.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.CenterContinuous);
        }

        var ScaleX = (double)CellWidthPx / (Image.Width);
        var ScaleY = (double)CellHeightPx / Image.Height;
        var TargetScale = Math.Min(ScaleX, ScaleY);

        var TargetWidthPx = Math.Floor(Image.Width * TargetScale - 4);
        var TargetHeightPx = Math.Floor(Image.Height * TargetScale - 20);

        //var ImageScaleWidth = Image.Width * TargetScale;
        //var ImageScaleHeight = Image.Height * TargetScale;
        //var OffsetX = (int)((CellWidthPx - ImageScaleWidth) / 2);
        //var OffsetY = (int)((CellHeightPx - ImageScaleHeight) / 2);

        Image
            .WithPlacement(XLPicturePlacement.Move)
            .WithSize((int)TargetWidthPx, (int)TargetHeightPx)
            .MoveTo(Cell, 2, 4);

        return this;
    }
    public CellTrack SetImage(byte[] Buffer)
    {
        var Stream = new MemoryStream(Buffer);
        SetImage(Stream);
        return this;
    }
    public CellTrack SetValue(XLCellValue Value)
    {
        Cell.SetValue(Value);
        return this;
    }
    public TValue GetValue<TValue>() => Cell.GetValue<TValue>();
    public XLCellValue GetValue() => Cell.Value;
    public double GetCellWidth() => Worksheet.Column(ColumnNumber).Width;
    public double GetCellHeight() => Worksheet.Row(RowNumber).Height;
    public double GetCellWidthPixels() => ConvertWidthToPixels(GetCellWidth());
    public double GetCellHeightPixels() => ConvertHeightToPixels(GetCellHeight());
}