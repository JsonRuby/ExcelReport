using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.Model;
using NPOI.XSSF.UserModel;

namespace ExcelReport
{
    internal static class ISheetExtend
    {
        /// <summary>
        /// 复制区间
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="sourceCells"></param>
        /// <param name="gapIndex">下移的行间距</param>
        public static void CopyRange(this ISheet sheet, IEnumerable<ICell> sourceCells, int gapIndex)
        {
            var cells = sourceCells as IList<ICell> ?? sourceCells.ToList();
            var minRow = cells.Min(t => t.RowIndex);
            var maxRow = cells.Max(t => t.RowIndex);
            var minCol = cells.Min(t => t.ColumnIndex);
            var maxCol = cells.Max(t => t.ColumnIndex);

            #region 填充数据和格式

            foreach (var sourceCell in cells)
            {
                if (sourceCell != null)
                {
                    var targetRow = sheet.GetRow(sourceCell.RowIndex + gapIndex) ??
                               sheet.CreateRow(sourceCell.RowIndex + gapIndex);
                    var targetCell = targetRow.GetCell(sourceCell.ColumnIndex) ??
                                     targetRow.CreateCell(sourceCell.ColumnIndex, sourceCell.CellType);
                    targetCell.CellStyle = sourceCell.CellStyle;
                    targetCell.ExtSetCellValue(sourceCell.ExtGetCellValue());
                }
            }

            #endregion

            #region 合并单元格

            for (var i = 0; i < sheet.NumMergedRegions; i++)
            {
                var mr = sheet.GetMergedRegion(i);
                if (mr.FirstRow >= minRow && mr.LastRow <= maxRow && mr.FirstColumn >= minCol && mr.LastColumn <= maxCol)
                {
                    var cellRange = new CellRangeAddress(mr.FirstRow + gapIndex, mr.LastRow + gapIndex, mr.FirstColumn,
                        mr.LastColumn);
                    sheet.AddMergedRegion(cellRange);
                }
            }

            #endregion
        }


        /// <summary>
        /// 获取Icell集合
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startPoint">其实坐标</param>
        /// <param name="endPoint">结束坐标</param>
        /// <param name="newIcell">是否产生新的ICell,默认:true</param>
        /// <returns></returns>
        public static List<ICell> GetICellCollection(this ISheet sheet, Point startPoint, Point endPoint,
            bool newIcell = true)
        {
            var rowIndexGap = 0;
            var colIndexGap = 0;
            var lastRowIndex = endPoint.X;
            var lastColumnIndex = endPoint.Y;

            //ICell.GetSpan can't access..
            if (sheet.GetRow(endPoint.X).GetCell(endPoint.Y).IsMergedCell(endPoint, out rowIndexGap, out colIndexGap))
            {
                lastRowIndex = endPoint.X + rowIndexGap;
                lastColumnIndex = endPoint.Y + colIndexGap;
            }

            var cellCollection = new List<ICell>();
            var tWorkbook = sheet.Workbook;
            var tSheet = newIcell ? tWorkbook.CreateSheet("TemplateWhatever") : tWorkbook.GetSheet("TemplateWhatever");

            for (var r = startPoint.X; r <= lastRowIndex; r++)
            {
                var row = sheet.GetRow(r);
                var tRow = tSheet.GetRow(r) ?? tSheet.CreateRow(r);
                for (var i = startPoint.Y; i <= lastColumnIndex; i++)
                {
                    var cell = row.GetCell(i);

                    if (cell != null)
                    {
                        if (newIcell)
                        {
                            var tCell = tRow.GetCell(cell.ColumnIndex) ?? tRow.CreateCell(cell.ColumnIndex, cell.CellType);
                            tCell.CellStyle = cell.CellStyle;
                            tCell.ExtSetCellValue(cell.ExtGetCellValue());
                            cellCollection.Add(tCell);
                        }
                        else
                        {
                            cellCollection.Add(cell);
                        }
                    }
                }
            }
            return cellCollection;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="cellPoint"></param>
        /// <param name="rowIndexGap"></param>
        /// <param name="colIndexGap"></param>
        /// <returns></returns>
        public static bool IsMergedCell(this ICell cell, Point cellPoint, out int rowIndexGap, out int colIndexGap)
        {
            var sheet = cell.Sheet;
            var regionsCount = sheet.NumMergedRegions;
            rowIndexGap = 0;
            colIndexGap = 0;
            for (var i = 0; i < regionsCount; i++)
            {
                var range = sheet.GetMergedRegion(i);
                if (range.FirstRow == cellPoint.X && range.FirstColumn == cellPoint.Y)
                {
                    rowIndexGap = range.LastRow - range.FirstRow;
                    colIndexGap = range.LastColumn - range.FirstColumn;
                    break;
                }
            }
            return sheet.GetRow(cellPoint.X).GetCell(cellPoint.Y).IsMergedCell;
        }
    }
}