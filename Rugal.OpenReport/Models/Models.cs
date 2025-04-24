
namespace Rugal.OpenReport.Models;

public class CopyMergeOption
{
    public int StartRow { get; set; }
    public int StartColumn { get; set; }
    public int EndRow { get; set; }
    public int EndColumn { get; set; }
    public int TargetRow { get; set; }
    public int TargetColumn { get; set; }
}
public class CopyRowOption
{
    public PasteTypes PasteType { get; set; } = PasteTypes.Overwrite;
    public PositionTypes PositionType { get; set; } = PositionTypes.Assign;
    public int TargetRow { get; set; }
}
public enum PositionTypes
{
    Assign,
    Move,
}
public enum PasteTypes
{
    Overwrite,
    InsertBefore,
    InsertAfter,
}