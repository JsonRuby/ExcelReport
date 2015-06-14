using System;
using System.Collections.Generic;
using System.Drawing;

namespace _01_订单示例
{
    public static class ItemLogic
    {
        public static List<Item> GetItems(int listCount = 3)
        {
            var r = new Random();
            var list = new List<Item>();
            for (var i = 0; i < listCount; i++)
            {
                list.Add(new Item
                {
                    ItemSeq = i + 1,
                    ItemImage = Image.FromFile("../../Image/C#高级编程.jpg").ToBuffer(),
                    ItemName = "C#高级编程#" + r.Next(1, 100),
                    ItemPrice = r.Next(40, 50),
                    ItemQty = new Random().Next(1, 10)
                });
            }
            return list;
        }
    }
}