using System.Collections.Generic;
using System.Drawing;
using NPOI.SS.UserModel;

namespace ExcelReport
{
    internal class ExtSheetFormatterContext : SheetFormatterContext
    {
        private IEnumerable<ICell> _pageTemplateCells;
        public int PageIndexCount { set; get; }
        private int _newPageCount = 0;

        public ExtSheetFormatterContext(
            ISheet sheet,
            IEnumerable<ElementFormatter> formatters,
            IEnumerable<ICell> pageTemplateCells,
            int pageIndexCount)
            : base(sheet, formatters)
        {
            _pageTemplateCells = pageTemplateCells;
            PageIndexCount = pageIndexCount;
        }

        /// <summary>
        /// 获取新增的页面的数量
        /// </summary>
        /// <returns></returns>
        public int GetNewPageCount()
        {
            return _newPageCount;
        }

        /// <summary>
        /// 判断startPoint到endPoint是否为空
        /// </summary>
        /// <param name="startPoint">起始坐标</param>
        /// <param name="endPoint">结束坐标</param>
        /// <returns></returns>
        private bool IsEmptySpan(Point startPoint, Point endPoint)
        {
            var isEmpty = true;
            // if (newPage) return true;
            for (var r = startPoint.X; r < endPoint.X; r++)
            {
                var tempRow = Sheet.GetRow(r);
                if (tempRow == null)
                {
                    isEmpty = false;
                    break;
                }
                for (var i = startPoint.Y; i < endPoint.Y; i++)
                {
                    var tempCell = tempRow.GetCell(i);
                    if (tempCell != null && tempCell.CellType != CellType.Blank)
                    {
                        if (tempCell.CellType == CellType.String)
                        {
                            isEmpty = string.IsNullOrEmpty(tempCell.StringCellValue);
                            break;
                        }
                        isEmpty = false;
                        break;
                    }
                }
            }
            return isEmpty;
        }


        /// 预处理位置及格式
        /// <param name="startItemPoint">Item起始坐标</param>
        /// <param name="endItemPoint">Item结束坐标</param>
        /// <param name="lastIncreaseIndex"></param>
        /// <param name="increaseIndex"></param>
        /// <param name="newPage"></param>
        public void PrepareItemSourceContainer(Point startItemPoint, Point endItemPoint, int lastIncreaseIndex,
            out int increaseIndex, out bool newPage)
        {
            var gapIndex = endItemPoint.X - startItemPoint.X + 1;
            var sourceItemStartPoint = new Point(startItemPoint.X - gapIndex, startItemPoint.Y);
            var sourceItemEndPoint = new Point(endItemPoint.X - gapIndex, endItemPoint.Y);
            var sourceCellCollection = Sheet.GetICellCollection(sourceItemStartPoint, sourceItemEndPoint,false);


            if (IsEmptySpan(startItemPoint, endItemPoint))
            {
                Sheet.CopyRange(sourceCellCollection, gapIndex);
                increaseIndex = lastIncreaseIndex + gapIndex;
                newPage = false;
            }
            else
            {
                Sheet.CopyRange(_pageTemplateCells, PageIndexCount * (_newPageCount + 1));

                _newPageCount += 1;
                increaseIndex = PageIndexCount * _newPageCount;
                newPage = true;
            }
        }
    }
}