using System;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Pathway.WPF.ImportExport.Logs.StyleSheetBuilder
{
    public interface IStyleSheetWorker
    {
        void Prepare(Stylesheet stylesheet);
        Stylesheet Stylesheet { get; set; }
        UInt32 IndexRefCellBaseEven { get; set; }
        UInt32 IndexRefCellBaseOdd { get; set; }
        UInt32 IndexRefCellBase { get; set; }
     
        UInt32 IndexRefCellHeaderBase { get; set; }

    }
}
