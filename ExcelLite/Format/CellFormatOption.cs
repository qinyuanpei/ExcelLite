using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLite.Format
{
   
       public enum CellFormatOption
       { 
       /// <summary>
       /// 使用默认样式进行重写
       /// </summary>
           OverWrite,
          /// <summary>
          /// 仅仅更新单元格内数值
          /// </summary>
           KeepStyle,
          /// <summary>
          /// 引用当前存在的样式
          /// </summary>
           ReferenceStyle      
       } 

     
}
