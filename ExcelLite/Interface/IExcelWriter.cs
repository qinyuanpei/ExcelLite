using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLite.Interface
{
    interface IExcelWriter
    {
        void AppendRow(object[] data);
        void AppendRow<T>(T entity) where T : class;
        void AppendRange(object[][] data);
        void AppendRange<T>(List<T> list) where T : class;
        void WriteRow(object[] data, int row, int column);
        void WriteRange(object[][] data,int row,int column);
        void WriteRange(object[][] data,string startCell);
        void WriteColumn(object[] data,int column,int row);
    
    
    }
}
