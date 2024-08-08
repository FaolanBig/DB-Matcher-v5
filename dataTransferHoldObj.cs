using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using Org.BouncyCastle.Tls.Crypto;
using System;
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
        protected int objectID;
        protected int primarySheet;
        protected int fromPrimary;
        protected int toPrimary;
        protected int secondarySheet;
        protected int fromSecondary;
        protected int toSecondary;
        protected int primaryColumn;
        protected int secondaryColumn;
        protected int secondaryfromColumn;
        protected int secondarytoColumn;
        protected int resultSheet;
        protected int resultColumn;

        protected IWorkbook workbook;

        public readonly string[] primaryContents;
        public readonly string[] secondaryContents;
        public readonly int[] resultRow; //second: row
        public dataTransferHoldObj(int objectID,
                                   IWorkbook workbook,
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

            this.workbook = workbook;
            this.resultRow = new int[this.toPrimary - this.fromPrimary];

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
        }
        public void Dispose()
        {
            ISheet tsheet = workbook.GetSheetAt(this.secondarySheet);
            ISheet resultSheet = workbook.GetSheetAt(this.resultSheet);
            
            for(int row = this.fromPrimary; row < this.toPrimary; row++)
            {
                IRow irow = tsheet.GetRow(this.resultRow[row]);
                IRow resultIrow = resultSheet.GetRow(row);

                for(int col = this.secondaryfromColumn; col < this.secondarytoColumn; col++)
                {
                    ICell cell = irow.GetCell(col);
                    ICell resultCell = resultIrow.GetCell(this.resultColumn + (col - this.secondaryfromColumn));

                    resultCell.SetCellValue(cell.ToString());
                }
            }
        }
    }
}
