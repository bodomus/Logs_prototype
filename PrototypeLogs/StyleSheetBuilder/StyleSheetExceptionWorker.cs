using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Pathway.WPF.ImportExport.Logs.CellBuilder;

namespace Pathway.WPF.ImportExport.Logs.StyleSheetBuilder
{
    public class StyleSheetExceptionWorker: IStyleSheetWorker
    {
        public void Prepare(Stylesheet stylesheet)
        {
            //Stylesheet = book.WorkbookStylesPart.Stylesheet;
            IndexRefCellBaseEven = AddDefautlStyleEven(ref stylesheet);
            IndexRefCellBaseOdd = AddDefautlStyleOdd(ref stylesheet);
            IndexRefCellBase = AddDefautlStyle(ref stylesheet);
            IndexRefCellHeaderBase = AddDefautlHeaderStyle(ref stylesheet);
            stylesheet.Save();
        }

        public Stylesheet Stylesheet { get; set; }
        public uint IndexRefCellBaseEven { get; set; }
        public uint IndexRefCellBaseOdd { get; set; }
        public uint IndexRefCellBase { get; set; }
        
        public uint IndexRefCellHeaderBase { get; set; }


        private UInt32 AddDefautlHeaderStyle(ref Stylesheet stylesheet)
        {
            var cellBuilder = new CellBuilder.CellBuilder(ref stylesheet);
            return cellBuilder
                .BuildFont(16, "000000")
                .BuildFill(PatternValues.Solid, "FFFFFF", "8c8b87")
                .BuildBorder(new BorderConfig
                {
                    BBorder = new BottomBorder()
                    {
                        Color = new Color() { Rgb = HexBinaryValue.FromString("fc92ed") },
                        Style = BorderStyleValues.Thick
                    },
                    TBorder = new TopBorder()
                    {
                        Color = new Color() { Rgb = HexBinaryValue.FromString("fc92ed") },
                        Style = BorderStyleValues.Thick
                    },
                    LBorder = new LeftBorder()
                    {
                        Color = new Color() { Rgb = HexBinaryValue.FromString("fc92ed") },
                        Style = BorderStyleValues.Thick
                    },
                    RBorder = new RightBorder()
                    {
                        Color = new Color() { Rgb = HexBinaryValue.FromString("fc92ed") },
                        Style = BorderStyleValues.Thick
                    }

                }).GetCellFormat(new Alignment()
                {
                    Horizontal = HorizontalAlignmentValues.Center,
                    Vertical = VerticalAlignmentValues.Center
                }).Key;
        }
        private UInt32 AddDefautlStyle(ref Stylesheet stylesheet)
        {
            var cellBuilder = new CellBuilder.CellBuilder(ref stylesheet);
            return cellBuilder
                .BuildFont(12, "FFFFFF")
                .BuildFill(PatternValues.Solid, "111111", "ffffff")
                .BuildBorder(new BorderConfig
                {
                    BBorder = new BottomBorder()
                    {
                        Color = new Color() { Rgb = HexBinaryValue.FromString("1122FF") },
                        Style = BorderStyleValues.Thin
                    },
                    TBorder = new TopBorder()
                    {
                        Color = new Color() { Rgb = HexBinaryValue.FromString("1122FF") },
                        Style = BorderStyleValues.Thin
                    },
                    LBorder = new LeftBorder()
                    {
                        Color = new Color() { Rgb = HexBinaryValue.FromString("1122FF") },
                        Style = BorderStyleValues.Thin
                    },
                    RBorder = new RightBorder()
                    {
                        Color = new Color() { Rgb = HexBinaryValue.FromString("1122FF") },
                        Style = BorderStyleValues.Thin
                    }

                }).GetCellFormat(new Alignment()
                {
                    Horizontal = HorizontalAlignmentValues.Left,
                    Vertical = VerticalAlignmentValues.Center
                }).Key;
        }
       
        private UInt32 AddDefautlStyleEven(ref Stylesheet stylesheet)
        {
            var cellBuilder = new CellBuilder.CellBuilder(ref stylesheet);
            return cellBuilder
                .BuildFont(12, "000000")
                .BuildFill(PatternValues.Solid, "ffffff", "cacfd9")
                .BuildBorder(new BorderConfig
                {
                    BBorder = new BottomBorder()
                    {
                        Color = new Color() { Rgb = HexBinaryValue.FromString("1122FF") },
                        Style = BorderStyleValues.Thin
                    },
                    TBorder = new TopBorder()
                    {
                        Color = new Color() { Rgb = HexBinaryValue.FromString("1122FF") },
                        Style = BorderStyleValues.Thin
                    },
                    LBorder = new LeftBorder()
                    {
                        Color = new Color() { Rgb = HexBinaryValue.FromString("1122FF") },
                        Style = BorderStyleValues.Thin
                    },
                    RBorder = new RightBorder()
                    {
                        Color = new Color() { Rgb = HexBinaryValue.FromString("1122FF") },
                        Style = BorderStyleValues.Thin
                    }

                }).GetCellFormat(new Alignment()
                {
                    Horizontal = HorizontalAlignmentValues.Left,
                    Vertical = VerticalAlignmentValues.Center
                }).Key;

        }

        private UInt32 AddDefautlStyleOdd(ref Stylesheet stylesheet)
        {
            var cellBuilder = new CellBuilder.CellBuilder(ref stylesheet);
            return cellBuilder
                .BuildFont(12, "000000")
                .BuildFill(PatternValues.Solid, "ffffff", "a2b8db")
                .BuildBorder(new BorderConfig
                {
                    BBorder = new BottomBorder()
                    {
                        Color = new Color() { Rgb = HexBinaryValue.FromString("1122FF") },
                        Style = BorderStyleValues.Thin
                    },
                    TBorder = new TopBorder()
                    {
                        Color = new Color() { Rgb = HexBinaryValue.FromString("1122FF") },
                        Style = BorderStyleValues.Thin
                    },
                    LBorder = new LeftBorder()
                    {
                        Color = new Color() { Rgb = HexBinaryValue.FromString("1122FF") },
                        Style = BorderStyleValues.Thin
                    },
                    RBorder = new RightBorder()
                    {
                        Color = new Color() { Rgb = HexBinaryValue.FromString("1122FF") },
                        Style = BorderStyleValues.Thin
                    }

                }).GetCellFormat(new Alignment()
                {
                    Horizontal = HorizontalAlignmentValues.Left,
                    Vertical = VerticalAlignmentValues.Center
                }).Key;
        }
    }
}