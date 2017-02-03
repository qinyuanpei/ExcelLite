using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Text.RegularExpressions;
using ExcelLite.Format;

namespace ExcelLite
{
   public static class SpreadsheetExtender
    {
       /// <summary>
       /// [扩展方法]根据名称获取指定Worksheet对象
       /// </summary>
       /// <param name="document"></param>
       /// <param name="sheetName"></param>
       public static Worksheet GetWorksheet(this SpreadsheetDocument document,string sheetName)
       {
       //查询指定工作表
           var sheet=document.WorkbookPart.Workbook.Descendants<Sheet>().
             FirstOrDefault(s =>s.Name==sheetName);
           if(sheet==null)
               throw new NullSheetException("Please make sure sheetname is correct!");

           //返回WorkSheet对象
           var worksheetPart=(WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);
           return worksheetPart.Worksheet;      
       }
       /// <summary>
       /// 
       /// </summary>
       /// <param name="document"></param>
       /// <param name="sheetIndex"></param>
       /// <returns></returns>
       public static Worksheet GetWorksheet(this SpreadsheetDocument document, int sheetIndex)
       {
           int sheetCount = document.GetWorksheetCount();
           if (sheetIndex < 0 || sheetIndex > sheetCount)
               throw new ArgumentOutOfRangeException("请确认输入的索引是否正确");
          //返回指定工作表
           Sheet sheet = document.WorkbookPart.Workbook.Descendants<Sheet>().ElementAt(sheetIndex);
           WorksheetPart workSheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);
           return workSheetPart.Worksheet;      
       }
       /// <summary>
       /// 返回当前WorkSheet的名称
       /// </summary>
       /// <param name="document"></param>
       /// <param name="worksheet"></param>
       /// <returns></returns>
       public static string GetWorksheetName(this SpreadsheetDocument document, Worksheet worksheet)
       { 
       //获取文档中全部工作表
           var sheets = document.WorkbookPart.Workbook.Descendants<Sheet>();
           var sheet = sheets.Where(s => ((WorksheetPart)document.WorkbookPart.GetPartById(s.Id)).Worksheet.Equals(worksheet)).FirstOrDefault();
           return sheet.Name; 
       }
       /// <summary>
       /// [扩展方法]返回当前文件样式表
       /// </summary>
       /// <param name="docunment"></param>
       /// <returns></returns>
       public static Stylesheet GetStylesheet(this SpreadsheetDocument docunment)
       {
           var workbookStylesPart = docunment.WorkbookPart.WorkbookStylesPart;
           if (workbookStylesPart == null)
               workbookStylesPart = docunment.WorkbookPart.AddNewPart<WorkbookStylesPart>();
           var styleSheet = workbookStylesPart.Stylesheet;
           if (styleSheet == null)
               styleSheet = new Stylesheet();
           return styleSheet;
        
       }
       public static int GetWorksheetCount(this SpreadsheetDocument document)
       {

           return document.WorkbookPart.Workbook.Descendants<Sheet>().Count(); 
       }
       /// <summary>
       /// [扩展方法]根据名称获取指定sheetdata对象
       /// </summary>
       /// <param name="document"></param>
       /// <returns></returns>
       public static SheetData GetSheetData(this SpreadsheetDocument document, string sheetName)
       { 
       //获取指定Worksheet
           var sheet = document.GetWorksheet(sheetName);
           if (sheet == null)
               throw new NullSheetException("请确认WorkSheet对象是否为null!");

           return sheet.GetFirstChild<SheetData>();
       
       }
       public static SheetData GetSheetData(this Worksheet worksheet)
       {
           if (worksheet == null)
               throw new NullSheetException("请确认WorkSheet对象是否为空！");
           return worksheet.GetFirstChild<SheetData>();
       
       }
       public static SharedStringTablePart GetSharedStringTable(this SpreadsheetDocument document)
       {
           if (document == null)
               throw new NullDocumentException("请确认SpreadsheetDocument是否为空！");
           return document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault(); 
       }
       /// <summary>
       /// [扩展方法]设 置指定单元格数值
       /// </summary>
       /// <param name="worksheet"></param>
       /// <param name="row"></param>
       /// <param name="column"></param>
       /// <param name="value"></param>
       public static void SetCellValue(this Worksheet worksheet, int row, int column, object value)
       {
           string cellAddress = ToCellAddress(row, column);
           worksheet.SetCellValue(cellAddress, value);      
       }
       /// <summary>
       /// [扩展方法]设置指定单元格数值 
       /// </summary>
       /// <param name="worksheet"></param>
       /// <param name="cellAddress"></param>
       /// <param name="value"></param>
       public static void SetCellValue(this Worksheet worksheet, string cellAddress, object value)
       {
           //根据单元格地址获取指定行
           SheetData sheetData = worksheet.GetFirstChild<SheetData>();
           uint rowIndex = ToRowIndex(cellAddress);
           Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
           //如果指定行不存在则先创建行
           if (row == null)
           {
               row = new Row { RowIndex = rowIndex };
               Cell cell = row.CreateCell(cellAddress);
               sheetData.Append(row);
           }
           else {
               //根据单元格地址获取指定列
               Cell cell = row.Elements<Cell>().FirstOrDefault(item => item.CellReference.Value == cellAddress);
               //如果指定列不存在则先创建列
               if (cell == null)
               {
                   cell=row.CreateCell(cellAddress);
                    cell.SetCellText(value);              
                              }
               else { cell.SetCellText(value); }   
           
           }

       }
       /// <summary>
       /// [扩展方法]设置单元格地址
       /// </summary>
       /// <param name="worksheet"></param>
       /// <param name="cell"></param>
       /// <param name="value"></param>
       /// <param name="styleIndex"></param>
       public static void SetCellValue(this Worksheet worksheet, Cell cell, object value, int styleIndex)
       {
           if (worksheet == null)            
               throw new NullSheetException("请确认Worksheet对象是否为Null！"); 
           //当单元格不存在时返回
           if (cell == null) return;
           //设置单元格地址
           string cellAddress = cell.CellReference.Value;
           //设置单元格数值
           worksheet.SetCellValue(cellAddress, value); 
       }
       /// <summary>
       /// [扩展方法]当Row为空引发异常返回值指定单元格数值，当行数和列数小于1时引发异常
       /// </summary>
       /// <param name="worksheet">Worksheet对象</param>
       /// <param name="cellAddress"></param>
       /// <param name="sharedTablePart"></param>
       /// <returns></returns>
       public static string GetCellValue(this Worksheet worksheet, string cellAddress, SharedStringTablePart sharedTablePart)
       {
         // 获取当前行索引
           uint rowIndex = ToRowIndex(cellAddress);
        //获取当前行
           var row = worksheet.Descendants<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
           if (row == null) return string.Empty;
           return worksheet.GetCellValue(cellAddress, sharedTablePart);
       
       }
        /// <summary>
       /// [扩展方法]当行数和列数小于1时引发异常
       /// </summary>
       /// <param name="worksheet">Worksheet对象</param>
       /// <param name="cellAddress"></param>
       /// <param name="sharedTablePart"></param>
       /// <returns></returns>
       /// <exception cref=""></exception>
       public static string GetCellValue(this Worksheet worksheet,int row ,int column,SharedStringTablePart sharedTablePart)
       {
          if(row<1 || column<1)throw new ArgumentException("Excel number of row and column starting with 1");
           //计算单元格地址
           string cellAddress=ToCellAddress(row,column);
           //返回当前单元格
           Row targetRow=worksheet.GetRow(row);
           if(targetRow==null)return string.Empty;
           Cell targetCell = targetRow.GetCell(cellAddress);
           if (targetCell == null) return string.Empty;
           return worksheet.GetCellValue(targetCell, sharedTablePart);
       }
       public static string GetCellValue(this Worksheet worksheet, Cell cell, SharedStringTablePart sharedTablePart)
       {
           if (cell == null) throw new NullReferenceException("Please confirm Cell object is null");
           //当单元格数值为空时返回string.Empty
           string cellValue = string.Empty;
           if (cell.ChildElements.Count <= 0)
               return cellValue;
           //获取Cell对象的InnerText
           string cellInnerText = cell.CellValue.InnerText;
           cellValue = cellInnerText;
           if(sharedTablePart==null)return cellInnerText;
           //获取SharedStringTable对象
           var sharedTable = sharedTablePart.SharedStringTable;
           try
           { //数据类型处理
               EnumValue<DocumentFormat.OpenXml.Spreadsheet.CellValues> cellType = cell.DataType;
               if (cellType != null)
               {
                   switch (cellType.Value)
                   {
                       case DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString:
                           int cellIndex = int.Parse(cellInnerText);
                           cellValue = sharedTable.ChildElements[cellIndex].InnerText;
                           break;
                       case DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean:
                           cellValue = (cellInnerText == "1") ? "TRUE" : "FALSE";
                           break;
                       case DocumentFormat.OpenXml.Spreadsheet.CellValues.Date:
                           cellValue = Convert.ToDateTime(cellInnerText).ToString(); ;
                           break;
                       case DocumentFormat.OpenXml.Spreadsheet.CellValues.Number:
                           cellValue = Convert.ToDecimal(cellInnerText).ToString(); ;
                           break;
                       default:
                           cellValue = cellInnerText;
                           break;
                   }
               }
           }
           catch 
           {

               cellValue = "N/A";
           }
           return cellValue;
       }
       public static Row GetRow(this Worksheet worksheet,int rowIndex)
       {
           if (rowIndex < 1)
               throw new ArgumentException("Please make sure RowIndex is in range of this worksheet!");
           //获取指定行
           var row = worksheet.GetSheetData().Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
           //返回指定行
           return row;     
       }
       public static int GetRowCount(this Worksheet worksheet)
       {
           if (worksheet == null) throw new NullSheetException("Please confirm whether the WorkSheet null!");
           //Table with merged cells can caculate in this way
           return worksheet.GetSheetData().Elements<Row>().Max(r => int.Parse(r.RowIndex.ToString())); 
       }
       public static Cell GetCell(this Row row, string cellAddress)
       {
           if (row == null)
               throw new CellAddressException("Make confirm whethere the current row exists!");
           if (string.IsNullOrEmpty(cellAddress)) throw new CellAddressException("Make sure the cell is empty or the format is correct!");
           //当目标单元格不存在时建的单元格
           var  cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value==cellAddress);
           //返回指定单元格
           return cell;       
       }

        public static Cell GetCell(this Row row, int cellIndex)
       {
           if (row == null)
               throw new NullReferenceException("Please confirm whethere the current row exists!");
           var cells = row.Elements<Cell>();
           if (cellIndex < 1)
               throw new CellAddressException("Please confirm whethere the column index within a predetermination!");
           int rowIndex = row.GetRowIndex();
           string cellAddress = ToCellAddress(rowIndex, cellIndex);
           var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == cellAddress);
           //返回指定单元格
           return cell;       
       }

        public static Cell GetCell(this Worksheet worksheet, string cellAddress)
        {
            int rowIndex = worksheet.GetRowIndex(cellAddress);
            int columnIndex = SpreadsheetExtender.ToCellIndex(cellAddress);
            return worksheet.GetCell(rowIndex, columnIndex);     
        }
        public static Cell GetCell(this Worksheet worksheet, int row, int column)
        {

            Row targetRow = worksheet.GetRow(row);
            return targetRow.GetCell(column);     
        }
        public static void SetCellText(this Cell cell, object value)
        {
            if (cell == value)
                throw new NullReferenceException("Please confirm Cell object is null!");
            //初始化Cellvalue
            if (cell.CellValue == null)
                cell.CellValue = new CellValue();
            CellType type = ToCellType(value);
            switch (type)
            { 
            
                case CellType.String:
                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                    cell.CellValue.Text = value.ToString();
                    break;

                case CellType.Boolean:
                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean;
                    cell.CellValue.Text =(((bool)value)==true) ?  "1":"0";
                    break;

                case CellType.Number:
                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
                    cell.CellValue.Text = value.ToString();
                    break;
                case CellType.Date://此处与源代码不同，疑似源代码笔误
                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Date;
                    cell.CellValue.Text = value.ToString();
                    break;
                case CellType.Formula:
                    string expression = value.ToString();
                    expression = expression.Substring(1, expression.Length - 1);
                    cell.SetCellFormula(expression);
                    break; 
            } 
        }
        public static void SetCellFormula(this Cell cell, string expression)
        {
            if (cell == null)
                throw new NullReferenceException("Please confirm Cell is null!");
            if(cell.CellFormula==null)
                cell.CellFormula=new CellFormula(){CalculateCell=true};
            cell.CellFormula.Text = expression; 
        }
        public static void SetCellStyle(this Cell cell, int styleIndex)
        {
            cell.StyleIndex = UInt32Value.FromUInt32((uint)styleIndex);       
        }
        public static Cell CreateCell(this Row row, string cellAddress)
        {
            var resultCell = new Cell { CellReference = cellAddress };
            var refCell = row.Elements<Cell>().FirstOrDefault(cell => string.Compare(cell.CellReference.Value, cellAddress, StringComparison.OrdinalIgnoreCase) > 0);
        if(refCell!=null)
            row.InsertBefore(resultCell,refCell);
        else
                row.Append(resultCell);
            return resultCell;        
        }
        public static int GetCellIndex(this Cell cell)
        {
            if (cell == null)
                throw new NullReferenceException("Please confirm Call object is null");
            return ToCellIndex(cell.CellReference.Value);      
        }
        public static int GetRowIndex(this Row row)
        {
            if (row == null)
                throw new NullReferenceException("Please confirm Row object is null!");
            return int.Parse(row.RowIndex.ToString());        
        }
        public static int GetCellIndex(this Worksheet workssheet, string cellAddress)
        {
            return ToCellIndex(cellAddress);      
        }
        public static int GetRowIndex(this Worksheet worksheet, string cellAddress)
        {
            return int.Parse(ToRowIndex(cellAddress).ToString());      
        }
        public static string ToCellAddress(int row,int column)
        { 
            //定义26个字母
            string[] characters = new string[] { "A","B","C","D","E","F","G","H","I","J","K","L","M",
                "N","O","P","Q","R","S","T","U","V","W","X","Y","Z"};
            
            //当行数和列数不符合要求时引发异常
            if (row < 1 || column < 1)
            {
                throw new ArgumentException("Excelnumber of rows and column from a start!");          
            }
            //的前列对应的名称

            string columnName=string.Empty;
            while (column % 26 > 26)
            {
                columnName += characters[(column / 26) - 1];
                column = column % 26;          
            }
            columnName += characters[column - 1];
            return columnName + row;
        }
        private static uint ToRowIndex(string cellAddress)
        {
            var regex = new Regex(@"[0-9]+");
            var match = regex.Match(cellAddress);
            return uint.Parse(match.Value);       
        }
        private static  int ToCellIndex(string cellAddrss)
        {
            string[] characters = new string[] { "A","B","C","D","E","F","G","H","I","J","K","L","M",
                "N","O","P","Q","R","S","T","U","V","W","X","Y","Z"};
          //查找一个数字开始的位置
            var first = cellAddrss.Select((value, index) => new { Value = value, Index = index }).First(item => char.IsDigit(item.Value));
           //获取当前单元格地址中的列名部分
            cellAddrss = cellAddrss.Substring(0, first.Index);
            int length = cellAddrss.Length;
            //当长度为1时直接返回其索引
            if (length == 1)
                return characters.IndexOf(cellAddrss) + 1;
            //计算当前列名对应的列号
            int reval = 0;
            for (int i = 0; i < length; i++)
            {
                string chars = cellAddrss.Substring(i, 1);
                int index = characters.IndexOf(chars) + 1;
                int power = (length - i - 1);
                reval += index * (int)Math.Pow(26, power); 
            }
            return reval;
        }
        private static int IndexOf(this string[] data, string target) 
        {
            if (!data.Contains(target) || data == null) return -1;
            //查找第一个匹配结果
            var first = data.Select((item, index) => new { Item = item, Index = index }).First(x => x.Item == target);
            //返回索引
            return first.Index; 
        }
        private static CellType ToCellType(object value)
        {
            Type type = value.GetType();
            switch (type.Name)
            {
                case "Int32":
                case "Int64":
                case "Single":
                case "Double":
                case "Decimal":
                    return CellType.Number;
                case "String":
                    //当字符串以"="开头时表示这是一个公式
                    //否则它会按照普通字符串处理
                    string text = value.ToString();
                    if (text.StartsWith("="))
                    {
                        return CellType.Formula;
                    }
                    else
                    {
                        return CellType.String;
                    }

                case "Boolean":
                    return CellType.Boolean;
                case "DateTime":
                    return CellType.Boolean;
                default:
                    return CellType.String;

            }
        }
       public static  int GetCellStyleIndex(Stylesheet styleSheet,object value)
       {
       
       var cellType=ToCellType(value);
           int styleIndex=0;
           switch(cellType)
           {
               case CellType.Date:
                   CellFormatBuilder.GetDefaultDateFormat(styleSheet,out styleIndex);break;
               case CellType.Number:
                   CellFormatBuilder.GetDefaultNumberFormat(styleSheet, out styleIndex);break;
               default: return styleIndex; 
           }

           return styleIndex;
       }

    }
   #region 异常定义
    class NullSheetException : Exception
   {

        public NullSheetException(string message) : base(message) { } 
   
   }
   class NullDocumentException : Exception
   {

       public NullDocumentException(string message)  :base(message){} 
   
   }
    class CellAddressException:Exception
    {
    
    
     public CellAddressException(string message)  :base(message){} 
    
    }
    public enum CellType
    { 
        String,
        Boolean,
        Date,
        Number,
        Shared,
        Formula
    
    
    }
#endregion
}
