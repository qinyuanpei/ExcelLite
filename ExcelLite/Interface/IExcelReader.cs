using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLite.Interface
{
    interface IExcelReader
    {
        object[] ReadColumn(int column, bool ignoredReader);
        object[] ReadRow(int Row);
        T ReadRow<T>(int row)where T:class;
          object[][] ReadRange(int startRow,int endRow);
          object[][] ReadRange(string startRow, string endRow);
         List<T> ReadRange<T>(int start,int end) where T:class;
         object[][] ReadAll(bool ignoredReader);
         List<T> ReadAll<T>(bool ignoredReader) where T : class;
    
    }
}
