using System.Collections.Generic;
using System.Drawing;
using NPOI.SS.UserModel;

namespace ExcelReport
{
    internal class ExtRepeaterFormatter<T> : RepeaterFormatter<T>
    {
        #region 成员字段

        private Point _startItemPoint;
        private Point _endItemPoint;
        private IEnumerable<T> _dataSource;
        private List<RepeaterCellInfo<T>> _cellInfoList = new List<RepeaterCellInfo<T>>();
        private int _gapIndex;
        private Point _pageStartPoint;
        private Point _pageEndPoint;
        private Point _pageInfoPoint;

        #endregion

        public ExtRepeaterFormatter(
            Point startTagCell,
            Point endTagCell,
            IEnumerable<T> dataSource,
            Point pageStartPoint,
            Point pageEndPoint,
            Point pageInfoPoint,
            params RepeaterCellInfo<T>[] cellInfos)
            : base(startTagCell, endTagCell, dataSource, cellInfos)
        {
            _startItemPoint = startTagCell;
            _endItemPoint = endTagCell;
            _dataSource = dataSource;
            _pageStartPoint = pageStartPoint;
            _pageEndPoint = pageEndPoint;
            _pageInfoPoint = pageInfoPoint;
            _gapIndex = endTagCell.X - startTagCell.X + 1;
            if (null != cellInfos && cellInfos.Length > 0)
            {
                _cellInfoList.AddRange(cellInfos);
            }
        }

        public override void Format(SheetFormatterContext context)
        {
            if (null == _cellInfoList || _cellInfoList.Count <= 0 || null == _dataSource)
            {
                return;
            }
            var templateCells = new List<ICell>();
            templateCells.AddRange(context.Sheet.GetICellCollection(_pageStartPoint, _pageEndPoint, true));
            var extContext = new ExtSheetFormatterContext(
                context.Sheet,
                context.Formatters,
                templateCells,
                _pageEndPoint.X - _pageStartPoint.X + 1);


            var increaseIndex = 0;
            var tmpIndex = 0;
            var itemCount = 0;
            var newPage = false;

            foreach (var itemSource in _dataSource)
            {
                if (itemCount++ > 0)
                {
                    tmpIndex = increaseIndex > 0
                        ? newPage
                            ? increaseIndex + _gapIndex
                            : increaseIndex
                        : _gapIndex;
                    var currentStartPoint = new Point(_startItemPoint.X + tmpIndex, _startItemPoint.Y);
                    var currentEndPoint = new Point(_endItemPoint.X + tmpIndex, _endItemPoint.Y);

                    extContext.PrepareItemSourceContainer(currentStartPoint, currentEndPoint, tmpIndex,
                        out increaseIndex, out newPage);

                    tmpIndex = newPage ? increaseIndex : tmpIndex;
                }

                #region 填充数据

                foreach (var cellInfo in _cellInfoList)
                {
                    var row = extContext.Sheet.GetRow(cellInfo.CellPoint.X + tmpIndex);
                    var cell = row.GetCell(cellInfo.CellPoint.Y);
                    SetCellValue(cell, cellInfo.DgSetValue(itemSource));
                }

                #endregion
            }

            #region 页脚

            var newPageCount = extContext.GetNewPageCount();
            for (var i = 0; i <= newPageCount; i++)
            {
                var pageInfoString = "Page:" + (i + 1) + "/" + (newPageCount + 1);
                var pageInfoPoint = new Point(_pageInfoPoint.X + i * extContext.PageIndexCount, _pageInfoPoint.Y);
                var pageInfoCell = extContext.Sheet.GetRow(pageInfoPoint.X).GetCell(pageInfoPoint.Y);
                pageInfoCell.ExtSetCellValue(pageInfoString);
            }

            #endregion

            #region 计算公式

            extContext.Sheet.ForceFormulaRecalculation = true;

            #endregion

            #region 删除Template Sheet
            extContext.Sheet.Workbook.RemoveSheetAt(extContext.Sheet.Workbook.GetSheetIndex("TemplateWhatever"));
            extContext.Sheet.Workbook.FirstVisibleTab = 0;
            #endregion

        }
    }
}