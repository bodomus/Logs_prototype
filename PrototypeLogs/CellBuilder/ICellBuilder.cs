using System.Collections.Generic;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Pathway.WPF.ImportExport.Logs.CellBuilder
{
    public interface ICellBuilder
    {
        ICellBuilder BuildFont(DoubleValue @size, string color);

        ICellBuilder BuildBorder(BorderConfig config);

        ICellBuilder BuildFill(EnumValue<PatternValues> patternType, string color, string backgroundColor);

        KeyValuePair<UInt32Value, CellFormat> GetCellFormat(Alignment alignment);
    }
}
