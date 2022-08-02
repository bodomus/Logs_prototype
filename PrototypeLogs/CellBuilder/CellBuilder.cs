using System;
using System.Collections.Generic;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Pathway.WPF.ImportExport.Logs.CellBuilder
{
    public class CellBuilder : ICellBuilder
    {
        private UInt32 _fontId = 0;
        private UInt32 _fillId = 0;
        private UInt32 _cellFormatId = 0;
        private UInt32 _borderId = 0;
        private readonly Stylesheet _stylesheet;
        public CellBuilder(ref Stylesheet stylesheet)
        {
            _stylesheet = stylesheet;
        }

        public ICellBuilder BuildFill(EnumValue<PatternValues> patternType, string color, string backgroundColor)
        {
            PatternFill pFill = new PatternFill() { PatternType = patternType };
            pFill.ForegroundColor = new ForegroundColor() { Rgb = HexBinaryValue.FromString(color) };
            pFill.BackgroundColor = new BackgroundColor() { Rgb = HexBinaryValue.FromString(backgroundColor) };
            _stylesheet.Fills.Append(new Fill() { PatternFill = pFill });
            _fillId = _stylesheet.Fills.Count++;
            return this;
        }

        public ICellBuilder BuildFont(DoubleValue @size, string color)
        {
            Font font = new Font(new FontSize() { Val = @size },
                new Color() { Rgb = HexBinaryValue.FromString(color) });
            _stylesheet.Fonts.AppendChild(font);
            _fontId = _stylesheet.Fonts.Count++;
            return this;
        }

        public KeyValuePair<UInt32Value, CellFormat> GetCellFormat(Alignment alignment)
        {
            CellFormat cellFormat = new CellFormat()
            {
                FontId = _fontId,
                FillId = _fillId,
                ApplyFill = true,
                BorderId = _borderId, 
                Alignment = alignment,
            };

            _stylesheet.CellFormats.AppendChild(cellFormat);
            var cellFormatId = _stylesheet.CellFormats.Count++;
            return new KeyValuePair<UInt32Value, CellFormat>(cellFormatId, cellFormat);
        }

        public ICellBuilder BuildBorder(BorderConfig config)
        {
            Border b = new Border()
            {
                LeftBorder = config.LBorder,
                BottomBorder = config.BBorder,
                TopBorder = config.TBorder,
                RightBorder = config.RBorder,
            };

            _stylesheet.Borders.Append(b);
            _borderId = _stylesheet.Borders.Count++;
            return this;
        }
    }
}
