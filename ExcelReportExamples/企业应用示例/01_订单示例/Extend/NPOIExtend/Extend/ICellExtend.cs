using System;
using System.Drawing;
using System.Linq;
using System.Reflection;
using NPOI.HSSF.Record;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelReport
{
    internal static class ICellExtend
    {
        /// 设置单元格值
        /// <param name="cell"></param>
        /// <param name="value"></param>
        public static void ExtSetCellValue(this ICell cell, object value)
        {
            if (null == cell)
            {
                return;
            }
            if (null == value)
            {
                cell.SetCellValue(string.Empty);
            }
            else
            {
                if (value.GetType().FullName.Equals("System.Byte[]"))
                {
                    var pictureIdx = cell.Sheet.Workbook.AddPicture((Byte[])value, PictureType.PNG);
                    var anchor = cell.Sheet.Workbook.GetCreationHelper().CreateClientAnchor();
                    var rowIndexGap = 0;
                    var columnIndexGap = 0;
                    anchor.Col1 = cell.ColumnIndex;
                    anchor.Col2 = cell.ColumnIndex +
                                  (cell.IsMergedCell(new Point(cell.RowIndex, cell.ColumnIndex), out rowIndexGap,
                                      out columnIndexGap)
                                      ? rowIndexGap
                                      : rowIndexGap);
                    anchor.Row1 = cell.RowIndex;
                    anchor.Row2 = cell.RowIndex +
                                  +(cell.IsMergedCell(new Point(cell.RowIndex, cell.ColumnIndex), out rowIndexGap,
                                      out columnIndexGap)
                                      ? columnIndexGap
                                      : columnIndexGap);

                    var patriarch = cell.Sheet.CreateDrawingPatriarch();
                    var pic = patriarch.CreatePicture(anchor, pictureIdx);
                }
                else if (cell.CellType == CellType.Formula)
                {
                    cell.SetCellFormula((string)value);
                }
                else
                {
                    var valueTypeCode = Type.GetTypeCode(value.GetType());
                    switch (valueTypeCode)
                    {
                        case TypeCode.String: //字符串类型
                            cell.SetCellValue(Convert.ToString(value));
                            break;

                        case TypeCode.DateTime: //日期类型
                            cell.SetCellValue(Convert.ToDateTime(value));
                            break;

                        case TypeCode.Boolean: //布尔型
                            cell.SetCellValue(Convert.ToBoolean(value));
                            break;

                        case TypeCode.Int16: //整型
                        case TypeCode.Int32:
                        case TypeCode.Int64:
                        case TypeCode.Byte:
                        case TypeCode.Single: //浮点型
                        case TypeCode.Double:
                        case TypeCode.UInt16: //无符号整型
                        case TypeCode.UInt32:
                        case TypeCode.UInt64:
                            cell.SetCellValue(Convert.ToDouble(value));
                            break;

                        default:
                            cell.SetCellValue(string.Empty);
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// 获取单元格值
        /// </summary>
        /// <param name="sourceCell"></param>
        public static object ExtGetCellValue(this ICell sourceCell)
        {
            if (sourceCell == null)
            {
                return null;
            }
            switch (sourceCell.CellType)
            {
                case CellType.Numeric:
                    return sourceCell.NumericCellValue;
                case CellType.String:
                    return sourceCell.StringCellValue;
                case CellType.Formula:
                    return sourceCell.CellFormula;
                case CellType.Blank:
                    return sourceCell.StringCellValue;
                case CellType.Boolean:
                    return sourceCell.BooleanCellValue;
                case CellType.Error:
                    return sourceCell.ErrorCellValue;
                default:
                    return null;
            }
        }
    }
}