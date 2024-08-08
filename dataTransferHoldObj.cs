using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
//using NPOI.XWPF.UserModel;
using Org.BouncyCastle.Tls.Crypto;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using static NPOI.HSSF.Util.HSSFColor;

namespace DB_Matcher_v5
{
    internal class dataTransferHoldObj : IDisposable
    {
        internal int objectID;
        internal int primarySheet;
        internal int fromPrimary;
        internal int toPrimary;
        internal int secondarySheet;
        internal int fromSecondary;
        internal int toSecondary;
        internal int primaryColumn;
        internal int secondaryColumn;
        internal int secondaryfromColumn;
        internal int secondarytoColumn;
        internal int resultSheet;
        internal int resultColumn;
        internal bool useStyles = false;

        protected IWorkbook workbook;
        protected ICellStyle[] styles;
        protected Dictionary<string, string> dictionary;

        internal string[] primaryContents;
        internal string[] secondaryContents;
        internal int[] resultRow; //first: row
        internal int[] ld_value; //first: row
        internal string[] matchingValue;
        public dataTransferHoldObj(int objectID,
                                   IWorkbook workbook,
                                   ICellStyle[] styles,
                                   Dictionary<string, string> dictionary,
                                   int primarySheet,
                                   int fromPrimary,
                                   int toPrimary,
                                   int secondarySheet,
                                   int fromSecondary,
                                   int toSecondary,
                                   int primaryColumn,
                                   int secondaryColumn,
                                   int secondaryfromColumn,
                                   int secondarytoColumn,
                                   int resultSheet,
                                   int resultColumn)
        {
            this.objectID = objectID;
            this.fromPrimary = fromPrimary;
            this.toPrimary = toPrimary;
            this.fromSecondary = fromSecondary;
            this.toSecondary = toSecondary;
            this.primarySheet = primarySheet;
            this.secondarySheet = secondarySheet;
            this.primaryColumn = primaryColumn;
            this.secondaryColumn = secondaryColumn;
            this.secondaryfromColumn = secondaryfromColumn;
            this.secondarytoColumn = secondarytoColumn;
            this.resultSheet = resultSheet;
            this.resultColumn = resultColumn;
            this.useStyles = SettingsAgent.GetSettingIsTrue("colorGradient");

            this.workbook = workbook;
            this.styles = styles;
            this.dictionary = dictionary;
            this.resultRow = new int[this.toPrimary - this.fromPrimary];
            this.ld_value = new int[this.toPrimary - this.fromPrimary];

            ToLog.Inf($"new dataTransferHoldObj initialized - objectID: {this.objectID} - parameters (sheet: from --> to): primary({this.primarySheet}: {fromPrimary} --> {toPrimary} - secondary({this.secondarySheet}: {fromSecondary} --> {toSecondary}))");

            ToLog.Inf($"loading primary content from sheet: {this.primarySheet}");

            this.primaryContents = new string[this.toPrimary - this.fromPrimary];
            ISheet tsheet = this.workbook.GetSheetAt(this.primarySheet);
            for (int row = this.fromPrimary; row < this.toPrimary; row++)
            {
                IRow currentRow = tsheet.GetRow(row);

                if (currentRow == null) { this.primaryContents[row] = ""; }
                else
                {
                    ICell cell = currentRow.GetCell(this.primaryColumn);

                    if (cell != null && cell.CellType == CellType.String) { this.primaryContents[row] = cell.ToString() ?? ""; }
                    else { this.primaryContents[row] = ""; }
                }
            }

            this.secondaryContents = new string[this.toSecondary - this.fromSecondary];
            tsheet = this.workbook.GetSheetAt(this.secondarySheet);
            for (int row = this.fromSecondary; row < this.toSecondary; row++)
            {
                IRow currentRow = tsheet.GetRow(row);

                if (currentRow == null) { this.secondaryContents[row] = ""; }
                else
                {
                    ICell cell = currentRow.GetCell(this.secondaryColumn);

                    if (cell != null && cell.CellType == CellType.String) { this.secondaryContents[row] = cell.ToString() ?? ""; }
                    else { this.secondaryContents[row] = ""; }
                }
            }
            
            if (VarHold.useDataFile)
            {
                foreach (var entry in dictionary)
                {
                    int foreachRuns = 0;
                    foreach (string item in this.primaryContents)
                    {
                        bool added = false;
                        if (item.Contains(entry.Key))
                        {
                            this.primaryContents[foreachRuns] = item.Replace(entry.Key, entry.Value);
                            added = true;
                        }
                        foreachRuns++;
                        if (added) { break; }
                    }
                }
            }

            if (VarHold.useDataFile)
            {
                foreach (var entry in dictionary)
                {
                    int foreachRuns = 0;
                    foreach (string item in this.secondaryContents)
                    {
                        bool added = false;
                        if (item.Contains(entry.Key))
                        {
                            this.primaryContents[foreachRuns] = item.Replace(entry.Key, entry.Value);
                            added = true;
                        }
                        foreachRuns++;
                        if (added) { break; }
                    }
                }
            }
        }
        public void Dispose()
        {
            ISheet tsheet = workbook.GetSheetAt(this.secondarySheet);
            ISheet resultSheet = workbook.GetSheetAt(this.resultSheet);
            
            for(int row = this.fromPrimary; row < this.toPrimary; row++)
            {
                IRow irow = tsheet.GetRow(this.resultRow[row]);
                IRow resultIrow = resultSheet.GetRow(row);

                ICell resultCell = resultIrow.CreateCell(this.resultColumn);
                resultCell.SetCellValue(this.matchingValue[row]);
                if (this.ld_value[row] < 10) { resultCell.CellStyle = this.styles[this.ld_value[row]]; }
                else { resultCell.CellStyle = this.styles[10]; }

                for (int col = this.secondaryfromColumn; col < this.secondarytoColumn; col++)
                {
                    ICell cell = irow.GetCell(col);
                    resultCell = resultIrow.CreateCell(this.resultColumn + 1 + (col - this.secondaryfromColumn));

                    resultCell.SetCellValue(cell.ToString() ?? "");
                }
            }
        }
        internal void setMatchingValue(int row, int ld, int hd, int jd)
        {
            this.matchingValue[row] = $"(ld|hd|jd): ({ld}|{hd}|{jd})";
        }
    }
}
