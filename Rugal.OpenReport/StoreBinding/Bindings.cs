using ClosedXML.Excel;
using Rugal.OpenReport.Services;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Rugal.OpenReport.StoreBinding;

public class BindingSet
{
    public List<BindingBase> Bindings { get; set; }
    public SheetTrack SheetTrack { get; set; }
    protected IXLWorksheet Worksheet => SheetTrack.Worksheet;
    public BindingSet(SheetTrack SheetTrack)
    {
        this.SheetTrack = SheetTrack;
        Bindings = [];
        InitBindingSet();
    }
    protected void InitBindingSet()
    {
        var Cells = Worksheet.CellsUsed()
            .Where(Item => IsBindingSet(Item.Value.ToString()))
            .ToArray();

        foreach (var Cell in Cells)
        {
            var BindingLine = Cell.Value.ToString();
            var Lines = BindingLine.Replace(" ", "").Split(';');
            foreach (var Binding in Lines)
            {
                if (Binding.StartsWith("${"))
                {
                    var ValuePath = Regex.Match(Binding, @"\$\{([^\]]+)\}").Groups[1].Value;
                    Bindings.Add(new ValueBinding(SheetTrack, Binding, Cell, ValuePath));
                    continue;
                }

                var Command = Regex.Match(Binding, @"\$\[([^\]]+)\]").Groups[1].Value.ToLower();
                switch (Command)
                {
                    case "for-row":
                        Bindings.Add(new ForRowBinding(SheetTrack, BindingLine));
                        break;
                    case "item":
                        break;
                    default:
                        break;
                }
            }
        }
    }

    #region Public Method
    public void WriteBinding()
    {
        foreach (var Binding in Bindings)
        {
            Binding.WriteBinding();
        }
    }
    #endregion

    #region Public Static Method
    public static bool IsBindingSet(string Value)
    {
        var IsBinding = Regex.IsMatch(Value, @"\$\[.+?\]") || Regex.IsMatch(Value, @"\$\{.+?\}");
        return IsBinding;
    }
    #endregion
}
public abstract class BindingBase
{
    public object Store => SheetTrack.Store;
    public SheetTrack SheetTrack { get; set; }
    public string BindingLine { get; set; }
    public BindingBase(SheetTrack SheetTrack, string BindingLine)
    {
        this.SheetTrack = SheetTrack;
        this.BindingLine = BindingLine;
    }
    public abstract void WriteBinding();
    public bool TryGetValue(string FullPath, out object Result)
    {
        var Paths = FullPath.Split('.');

        var TargetStore = Store;
        foreach (var Path in Paths)
        {
            var GetProperty = TargetStore
                .GetType()
                .GetProperty(Path, BindingFlags.Instance | BindingFlags.Public);

            if (GetProperty is null)
            {
                Result = null;
                return false;
            }

            TargetStore = GetProperty.GetValue(TargetStore);
        }

        Result = TargetStore;
        return TargetStore is not null;
    }
    public bool TryGetValue<T>(string FullPath, out T Result)
    {
        if (!TryGetValue(FullPath, out var TryGetResult))
        {
            Result = default;
            return false;
        }

        if (TryGetResult is T TResult)
        {
            Result = TResult;
            return true;
        }

        Result = (T)TryGetResult;
        return true;
    }
}
public class ValueBinding : BindingBase
{
    public string ValuePath { get; protected set; }
    public IXLCell Cell { get; protected set; }
    public ValueBinding(SheetTrack SheetTrack, string BindingLine, IXLCell Cell, string ValuePath) : base(SheetTrack, BindingLine)
    {
        this.Cell = Cell;
        this.ValuePath = ValuePath;
    }
    public override void WriteBinding()
    {
        if (TryGetValue(ValuePath, out var ValueResult))
        {
            Cell.Value = ValueResult.ToString();
            return;
        }

        Cell.Value = $"[ValueBinding]: path [{ValuePath}] binding error";
    }
}
public class ForRowBinding : BindingBase
{
    public ForRowBinding(SheetTrack SheetTrack, string BindingLine) : base(SheetTrack, BindingLine)
    {

    }
    public override void WriteBinding()
    {

    }
}
public class ItemBinding : BindingBase
{
    public ItemBinding(SheetTrack SheetTrack, string BindingLine) : base(SheetTrack, BindingLine)
    {
    }
    public override void WriteBinding()
    {

    }
}