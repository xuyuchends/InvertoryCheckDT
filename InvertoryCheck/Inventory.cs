using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace InvertoryCheck
{
   public  class Inventory
    {
       /// <summary>
       /// 款号
       /// </summary>
       public string KuanHao { get; set; }
       /// <summary>
       /// 可用库存
       /// </summary>
       public int KeYongKuCun { get; set; }
       /// <summary>
       /// 实际库存
       /// </summary>
       public int ShiJiKuCun { get; set; }

    }
}
