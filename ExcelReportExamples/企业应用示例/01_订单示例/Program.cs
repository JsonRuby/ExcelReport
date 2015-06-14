using System;
using System.Collections.Generic;
using System.Linq;
using ExcelReport;

namespace _01_订单示例
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var collection = new ParameterCollection();
            collection.Load(@"../../Template/Template.xml");
            var lst = ItemLogic.GetItems(100);
            var formatters = new List<ElementFormatter>
            {
                new PartFormatter(collection["Sheet1", "PageStart"], "PageStart", ""),
                new PartFormatter(collection["Sheet1", "PageEnd"], "PageEnd", ""),
                new CellFormatter(collection["Sheet1", "ItemStart"], ""),
                new CellFormatter(collection["Sheet1", "ItemEnd"], ""),
                new CellFormatter(collection["Sheet1", "OrderNo"], "115052076000641601"),
                new CellFormatter(collection["Sheet1", "OrderDate"], DateTime.Now),
                new CellFormatter(collection["Sheet1", "ShipToAddress"], "A省B市C区D大道E号"),
                new PartFormatter(collection["Sheet1", "TrackingNo"], "TrackingNo", "199635102932"),
                new CellFormatter(collection["Sheet1", "OrderType"], "CTO"),
                new CellFormatter(collection["Sheet1", "PaymentType"], "FAC"),
                new CellFormatter(collection["Sheet1", "Remark"], "三方贸易,美元"),
                new CellFormatter(collection["Sheet1", "LegalAmount"], lst.Sum(t => t.ItemPrice*t.ItemQty)),
                new CellFormatter(collection["Sheet1", "Discount"], 50),
                new CellFormatter(collection["Sheet1", "TelPhone"], "138888888888"),

                //将RepeatFormatter信息放在最后.否则还要扩展其他formatter..
                //或者在ExcelReport中新增sheetFormatterContext._increaseIndex setter
                new ExtRepeaterFormatter<Item>(
                    collection["Sheet1", "ItemStart"],
                    collection["Sheet1", "ItemEnd"],
                    lst,
                    collection["Sheet1", "PageStart"],
                    collection["Sheet1", "PageEnd"],
                    collection["Sheet1", "PageInfo"],
                    new RepeaterCellInfo<Item>(collection["Sheet1", "ItemSeq"], t => t.ItemSeq),
                    new RepeaterCellInfo<Item>(collection["Sheet1", "ItemImage"], t => t.ItemImage),
                    new RepeaterCellInfo<Item>(collection["Sheet1", "ItemName"], t => t.ItemName),
                    new RepeaterCellInfo<Item>(collection["Sheet1", "ItemPrice"], t => t.ItemPrice),
                    new RepeaterCellInfo<Item>(collection["Sheet1", "ItemQty"], t => t.ItemQty)
                    ),
            };
            ExportHelper.ExportToLocal(@"../../Template/Template.xls",
                "result.xls",//bin/debug/
                new SheetFormatterContainer("Sheet1", formatters)
                );
        }
    }
}