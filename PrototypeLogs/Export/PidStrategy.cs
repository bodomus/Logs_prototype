using DocumentFormat.OpenXml.Spreadsheet;
using Pathway.WPF.Models;
using PrototypeLogs.Domain;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static PrototypeLogs.Export.LogsOpenXML;

namespace PrototypeLogs.Export
{
    public class PidStrategy : BaseStrategy, IExportExcelStrategy
    {
        private Dictionary<int, string> _sheetHeader
        {
            get
            {
                return new Dictionary<int, string>() {
                    {1, "TimeStamp"},
                    {2, "P"},
                    {3, "I"},
                    {4, "D"},
                    {5, "Error"},
                    {6, "SetPoint"},
                    {7, "OldSetPoint"},
                    {8, "Temperature1"},
                    {9, "Temperature2"},
                    {10, "DAC"},
                    {11, "RealTemperature1"},
                    {12, "RealTemperature2"},
                    {13, "WaterTemp"},
                    {14, "PCB"},
                    {15, "Heatsink1Temp"},
                    {16, "Heatsink2Temp"},
                    {17, "TEC"},
                    {18, "SensorMismatch"},
                    };
            }
        }

        private List<ColumnsPreference> _colunmPreferences
        {
            get
            {
                return new List<ColumnsPreference>() {
                    new ColumnsPreference{ 
                        Min = 1, Max = 1, Width = 30
                    },
                    new ColumnsPreference{
                        Min = 2, Max = 2, Width = 30
                    },
                    new ColumnsPreference{
                        Min = 3, Max = 3, Width = 30
                    },
                    new ColumnsPreference{
                        Min = 4, Max = 4, Width = 30
                    },
                    new ColumnsPreference{
                        Min = 5, Max = 5, Width = 30
                    },
                    new ColumnsPreference{
                        Min = 6, Max = 6, Width = 30
                    },
                    new ColumnsPreference{
                        Min = 7, Max = 7, Width = 30
                    },
                    new ColumnsPreference{
                        Min = 8, Max = 8, Width = 30
                    },
                    new ColumnsPreference{
                        Min = 9, Max = 9, Width = 30
                    },
                    new ColumnsPreference{
                        Min = 10, Max = 10, Width = 30
                    },
                    new ColumnsPreference{
                        Min = 11, Max = 11, Width = 30
                    },
                    new ColumnsPreference{
                        Min = 12, Max = 12, Width = 30
                    },
                    new ColumnsPreference{
                        Min = 13, Max = 13, Width = 30
                    },
                    new ColumnsPreference{
                        Min = 14, Max = 14, Width = 30
                    },
                    new ColumnsPreference{
                        Min = 15, Max = 15, Width = 30
                    },
                    new ColumnsPreference{
                        Min = 16, Max = 16, Width = 30
                    },
                    new ColumnsPreference{
                        Min = 17, Max = 17, Width = 30
                    },
                    new ColumnsPreference{
                        Min = 18, Max = 18, Width = 30
                    },
                    };
            }
        }
        public PidStrategy(string excelFileName, string logFileName, uint strategyIndex) : base(excelFileName, logFileName, strategyIndex)
        {
        }

        private string GetSheetName()
        {
            return Path.GetFileNameWithoutExtension(_logFileName);
        }
        public void DoWork()
        {
            rowIdx = 1;
            var strings = ReadFile(new LogFileTextReader(_logFileName)).ToList<string>();
            var sheetName = GetSheetName();
            var excel = new LogsOpenXML(_excelFileName, sheetName, _strategyIndex,  true, _colunmPreferences, false);
            excel.AddHeader(_sheetHeader);
            strings.RemoveAt(0);
            foreach (var s in strings)
            {
                var p = GetPidItem(s);
                rowIdx++;
                Row row = excel.SheetData.AppendChild(new Row() { RowIndex = rowIdx });
                excel.InsertCell(row, p.TimeStamp, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                excel.InsertCell(row, p.P, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                excel.InsertCell(row, p.I, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                excel.InsertCell(row, p.D, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                excel.InsertCell(row, p.Error, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                excel.InsertCell(row, p.SetPoint, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                excel.InsertCell(row, p.OldSetPoint, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                excel.InsertCell(row, p.Temperature1, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                excel.InsertCell(row, p.Temperature2, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                excel.InsertCell(row, p.DAC, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                excel.InsertCell(row, p.RealTemperature1, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                excel.InsertCell(row, p.RealTemperature2, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                excel.InsertCell(row, p.WaterTemp, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                excel.InsertCell(row, p.PCB, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                excel.InsertCell(row, p.Heatsink1Temp, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                excel.InsertCell(row, p.Heatsink2Temp, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                excel.InsertCell(row, p.TEC, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
                excel.InsertCell(row, p.SensorMismatch, CellValues.String, ExcelConstants.BOLDINDEXSTYLE);
            }
            excel.Close();
        }

        private PIDLog GetPidItem(string input)
        {
            if (string.IsNullOrEmpty(input))
                return null;
            var ar = input.Split(',');

            PIDLog result = new PIDLog()
            {
                TimeStamp = ar[0],
                P = ar[1],
                I = ar[2],
                D = ar[3],
                Error = ar[4],
                SetPoint = ar[5],
                OldSetPoint = ar[6],
                Temperature1 = ar[7],
                Temperature2 = ar[8],
                DAC = ar[9],
                RealTemperature1 = ar[10],
                RealTemperature2 = ar[11],
                WaterTemp = ar[12],
                PCB = ar[13],
                Heatsink1Temp = ar[14],
                Heatsink2Temp = ar[15],
                TEC = ar[16],
                SensorMismatch = ar[17],
            };
            return result;


            ////TimeStamp,P,I,D,Error,SetPoint,OldSetPoint,Temperature1,Temperature2,DAC,RealTemperature1,RealTemperature2,WaterTemp,PCB,Heatsink1Temp,Heatsink2Temp,TEC,SensorMismatch,
            //Regex expression = new Regex(@"(?<TimeStamp>.*),(?<P>.*),(?<I>.*),(?<D>.*),(?<Error>.*)
            //    ,(?<SetPoint>.*),(?<OldSetPoint>.*),(?<Temperature1>.*),(?<Temperature2>.*),(?<DAC>.*),(?<RealTemperature1>.*)
            //    ,(?<RealTemperature2>.*),(?<WaterTemp>.*),(?<PCB>.*),(?<Heatsink1Temp>.*),(?<Heatsink2Temp>.*),(?<TEC>.*),(?<SensorMismatch>.*)$");

            //Match match = expression.Match(input);

            //if (match.Success)
            //{
            //    return new PIDLog()
            //    {
            //        D = match.Groups["D"].Value,
            //    };
            //}
            //return null;
        }
    }
}
