using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.Data;
using System.Reflection;
using System.IO;
using System.Threading.Tasks;
using ExcelLite.Interface;
using ExcelLite.Config;
using ExcelLite.Format;

namespace ExcelLite.Helper
{
    public class ExcelHelper : IExcelReader, IExcelWriter, IDisposable
    {
        #region 字段定义
        private SpreadsheetDocument document;
        private DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet;
        //private Microsoft.Office.Interop.Excel.Worksheet Excelworksheet;
        private SheetData sheetData;
        private Stylesheet styleSheet;
        private int header;
        public int Header
        {
            get { return header; }
            set { header = value; }
        }
        private CellFormatOption cellFormatOption;
        public CellFormatOption CellFormatOption
        {
            get { return cellFormatOption; }
            set { cellFormatOption = value; }
        }
        #endregion
        #region 构造函数
        public ExcelHelper(string fileName, bool editable = true, int header = 1)
        {
            if (string.IsNullOrEmpty(fileName))
                throw new Exception("Check if current filename is empty!");
            if (!File.Exists(fileName))
                throw new Exception("Please make sure current file exists!");
            document = SpreadsheetDocument.Open(fileName, editable);
            worksheet = document.GetWorksheet(0);
            sheetData = worksheet.GetSheetData();
            styleSheet = document.GetStylesheet();
            this.Header = header;
            this.cellFormatOption = CellFormatOption.OverWrite;
        }
        #endregion
        public void SetHeader(int header)
        {
            this.header = header;
        }
        public void SetActiveSheet(int index)
        {
            this.Save();
            this.worksheet = document.GetWorksheet(index);
            this.sheetData = worksheet.GetSheetData();
        }
        public void SetActiveSheet(string sheetName)
        {
            this.Save();
            this.worksheet = document.GetWorksheet(sheetName);
            this.sheetData = worksheet.GetSheetData();
        }
        private void SetCellValue(string cellAddress, object value)
        {
            worksheet.SetCellValue(cellAddress, value);
            Cell cell = worksheet.GetCell(cellAddress);
            SetStyleIndexByCellFormatOption(cell, value, CellFormatOption);
        }
        public string GetCellValue(int row, int column)
        {
            var sharedTablePart = document.GetSharedStringTable();
            return worksheet.GetCellValue(row, column, sharedTablePart);
        }
        public string GetCellValue(string cellAddress)
        {
            var sharedTablePart = document.GetSharedStringTable();
            return worksheet.GetCellValue(cellAddress, sharedTablePart);
        }
        public int GetRowCount()
        {
            return worksheet.GetRowCount();

        }
        public void Close()
        {
            document.Close();
            document.Dispose();
        }
        public void Save()
        {
            worksheet.Save();

        }
        public void Dispose()
        {
            this.Save(); this.Close();

        }
        public void AppendRow(object[] data)
        {
            if (data == null || data.Length <= 0)

                throw new ArgumentNullException("Please confirm whether current input paraameter is null!");
            int rowIndex = worksheet.GetRowCount() + 1;

            Row row = new Row() { RowIndex = (uint)rowIndex };
            WriteRow(data, rowIndex);

        }
        public void AppendRow<T>(T entity) where T : class
        {

            if (entity == null) throw new ArgumentNullException("Please confirm whether current entity is null!");
            var cellMapping = GetHeaderCellMapping(typeof(T));
            var rowIndex = worksheet.GetRowCount() + 1;
            Row row = ConvertToRow<T>(entity, rowIndex, cellMapping);
            sheetData.Append(row);
        }
        public void AppendRange(object[][] data)
        {
            if (data == null || data.Length <= 0)
                throw new ArgumentNullException("Please confirm whether input datas is empty!");
            int rowIndex = worksheet.GetRowCount();
            Row[] rows = new Row[data.Length];
            for (int i = 0; i < data.Length; i++)
            {
                rowIndex += 1;
                rows[i] = new Row { RowIndex = (uint)rowIndex };

                for (int j = 0; j < data[i].Length; j++)
                {
                    string cellAddress = SpreadsheetExtender.ToCellAddress(rowIndex, (j + 1));
                    Cell cell = rows[i].CreateCell(cellAddress);
                    cell.SetCellText(data[i][j]);
                    int styleIndex = SpreadsheetExtender.GetCellStyleIndex(styleSheet, data[i][j]);
                    cell.SetCellStyle(styleIndex);
                }
            }
            sheetData.Append(rows);
        }
        public void AppendRange<T>(List<T> list) where T : class
        {
            if (list == null || list.Count == 0)
                throw new ArgumentNullException("Please make sure data wanted to Append  in list is exist!");
            var cellMapping = GetHeaderCellMapping(typeof(T));
            Row[] rows = new Row[list.Count];
            int rowIndex = worksheet.GetRowCount();
            for (int i = 0; i < list.Count; i++)
            {
                rowIndex += 1;
                rows[i] = ConvertToRow<T>(list[i], rowIndex, cellMapping);

            }
            sheetData.Append(rows);

        }

        public void WriteRange(object[][] data, string startCell)
        {
            int startRowIndex = worksheet.GetRowIndex(startCell);
            int startCellIndex = worksheet.GetRowIndex(startCell);
            List<Row> newRows = new List<Row>();
            for (int i = 0; i < data.Length; i++)
            {
                int rowIndex = startRowIndex + 1;
                Row row = worksheet.GetRow(rowIndex);
                if (row == null)
                {
                    row = new Row() { RowIndex = (uint)rowIndex };
                    newRows.Add(row);
                }
                for (int j = 0; j < data[i].Length; j++)
                {
                    int cellIndex = startCellIndex + j;
                    string cellAddress = SpreadsheetExtender.ToCellAddress(rowIndex, cellIndex);
                    Cell resultCell = row.GetCell(cellAddress);
                    if (resultCell == null)
                        resultCell = row.CreateCell(cellAddress);
                    resultCell.SetCellText(data[i][j]);
                    int styleIndex = SpreadsheetExtender.GetCellStyleIndex(styleSheet, data[i][j]);
                    resultCell.SetCellStyle(styleIndex);
                }
                sheetData.Append(newRows.ToArray());
            }
        }
        public void WriteRange(object[][] data, int row, int column)
        {
            string cellAddress = SpreadsheetExtender.ToCellAddress(row, column);
            WriteRange(data, cellAddress);
        }
        public void WriteColumn(object[] data, int column, int row = 1)
        {
            object[][] table = new object[data.Length][];
            for (int i = 0; i < table.Length; i++)
            {
                table[i] = new object[] { data[i] };
            }
            string cellAddress = SpreadsheetExtender.ToCellAddress(row, column);
        }

        public void WriteRow(object[] data, int row, int column = 1)
        {
            object[][] table = new object[1][];
            table[0] = data;
            string cellAddress = SpreadsheetExtender.ToCellAddress(row, column);
            WriteRange(table, cellAddress);
        }
        #region 读取方法

        public T ReadRow<T>(int row) where T : class
        {
            if (row < 1 | row > worksheet.GetRowCount())
                throw new ArgumentException("Make sure you enter the parameter within the specified range!");
            var sharedTablePart = document.GetSharedStringTable();
            var cellMapping = GetHeaderCellMapping(typeof(T));
            ConstructorInfo constructInfo = typeof(T).GetConstructor(Type.EmptyTypes);
            return ConvertToGeneric<T>(row, sharedTablePart, cellMapping, constructInfo);

        }
        public List<T> ReadRange<T>(int start, int end) where T : class
        {
            int count = worksheet.GetRowCount();
            if (start < 1 || end > count)
                throw new ArgumentException("Make sure you enter the correct parameter!");
            if (start > end)
                throw new ArgumentException("Make sure the end of the number of rows is greater than number of begin!");
            var sharedTablePart = document.GetSharedStringTable();
            var cellMapping = GetHeaderCellMapping(typeof(T));
            ConstructorInfo constructInfo = typeof(T).GetConstructor(Type.EmptyTypes);

            List<T> data = new List<T>();
            for (int i = start; i < end; i++)
            {
                data.Add(ConvertToGeneric<T>(i, sharedTablePart, cellMapping, constructInfo));

            }
            return data;
        }

        public object[] ReadRow(int rowIndex)
        {
            var row = worksheet.GetRow(header);
            if (row == null) throw new ArgumentException("We are unable to find the row with specified in rows!");
            var cells = row.Elements<Cell>();
            var sharedTablePart = document.GetSharedStringTable();
            int length = cells.Count();
            object[] data = new object[length];
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = worksheet.GetCellValue(rowIndex, (i + 1), sharedTablePart);
            }
            return data;
        }

        public object[] ReadColumn(int columnIndex, bool ignoreHeader = true)
        {
            string startCell = string.Empty;
            string endCell = string.Empty;
            object[] resultArray;
            if (!ignoreHeader)
            {

                startCell = SpreadsheetExtender.ToCellAddress(header, columnIndex);
                endCell = SpreadsheetExtender.ToCellAddress(worksheet.GetRowCount(), columnIndex);
            }
            var table = ReadRange(startCell, endCell);
            resultArray = new object[table.Length];
            for (int i = 0; i < table.Length; i++)
            {
                resultArray[i] = table[i][0];
            }
            return resultArray;
        }
        public object[][] ReadAll(bool ignorHeader)
        {
            int end = worksheet.GetRowCount();
            if (!ignorHeader)
            {
                return ReadRange(header, end);
            }
            return ReadRange(header + 1, end);
        }

        public List<T> ReadAll<T>(bool ignoreHeader = true) where T : class
        {

            int end = worksheet.GetRowCount();
            if (!ignoreHeader)
                return ReadRange<T>(header + 1, end);

            return ReadRange<T>(header + 1, end);
        }

        public object[][] ReadRange(int startRow, int endRow)
        {
            int count = worksheet.GetRowCount();
            if (startRow < 1 || endRow > count)
                throw new ArgumentException("Make sure you enter the correct parameters!");
            if (startRow >= endRow)
                throw new ArgumentException("Make sure the end of the number of rows is greater than begin!");
            object[][] data = new object[endRow - startRow + 1][];
            for (int i = startRow; i < endRow; i++)
            {
                data[i = startRow] = ReadRow(i);
            }
            return data;
        }

        public object[][] ReadRange(string startCell, string endCell)
        {
            int startRow = worksheet.GetRowIndex(startCell);
            int startColumn = worksheet.GetCellIndex(startCell);
            int endRow = worksheet.GetRowIndex(endCell);
            int endColumn = worksheet.GetCellIndex(endCell);
            if (startRow > endRow) throw new ArgumentException("StartCellAddress is not allowed to large than EndCellAddress!");
            object[][] data = new object[endRow = startRow + 1][];
            for (int i = startRow; i < endColumn; i++)
            {
                data[i - startRow] = new object[endColumn - startColumn + 1];
                for (int j = startColumn; j < endColumn + 1; j++)
                {
                    data[i - startRow][j - startColumn] = GetCellValue(i, j);
                }

            }
            return data;
        }
        private Row ConvertToRow<T>(T entity, int rowIndex, Dictionary<string, Cell> cellMapping)
        {
            if (entity == null)
            {
                throw new ArgumentException("Please confirm whether curren entity is null!");

            }

            Row row = new Row() { RowIndex = (uint)rowIndex };
            Type type = entity.GetType();
            PropertyInfo[] properties = type.GetProperties();
            foreach (PropertyInfo property in properties)
            {
                object[] attrs = property.GetCustomAttributes(typeof(MappingAttribute), true);
                if (attrs == null || attrs.Length <= 0) continue;
                MappingAttribute attr = (MappingAttribute)attrs[0];
                object value = property.FastGetValue(entity);
                if (value == null) value = string.Empty;
                string cellAddress = GetCellAddressByMapping(rowIndex, attr, cellMapping);
                Cell resultCell = row.CreateCell(cellAddress);
                resultCell.SetCellText(value);
                SetStyleIndexByCellFormatOption(resultCell, value, cellFormatOption, attr);
            }
            return row;
        }

        private T ConvertToGeneric<T>(int rowIndex, SharedStringTablePart sharedTablePart, Dictionary<string, Cell> cellMapping, ConstructorInfo constructorInfo)
        {
            T target = (T)constructorInfo.Invoke(null);
            Type type = target.GetType();
            PropertyInfo[] properties = type.GetProperties();
            foreach (PropertyInfo property in properties)
            {
                object[] attrs = property.GetCustomAttributes(typeof(MappingAttribute), true);
                MappingAttribute attr = (MappingAttribute)attrs[0];
                string cellValue = GetCellValueWithMapping(rowIndex, attr, cellMapping, sharedTablePart);
                object propertyValue = ParserFromString(property.PropertyType, cellValue);
            }
            return target;
        }

        private Dictionary<string, Cell> GetHeaderCellMapping(Type type)
        {
            Dictionary<string, Cell> results = new Dictionary<string, Cell>();
            var sharedTablePart = document.GetSharedStringTable();
            var headerCell = worksheet.GetRow(header).Elements<Cell>();
            if (headerCell == null)
                throw new NullReferenceException("Check that the header information exist!");
            PropertyInfo[] properties = type.GetProperties();
            foreach (PropertyInfo property in properties)
            {
                foreach (Cell cell in headerCell)
                {
                    object[] attrs = property.GetCustomAttributes(typeof(MappingAttribute), true);
                    if (attrs == null || attrs.Length <= 0) continue;
                    MappingAttribute attr = (MappingAttribute)attrs[0];
                    if (attr.MappingKey == worksheet.GetCellValue(cell, sharedTablePart))
                    {
                        results.Add(attr.MappingKey, cell); break;
                    }
                }
            }
            return results;
        }
        private string GetCellAddressByMapping(int rowIndex, MappingAttribute attribute, Dictionary<string, Cell> cellMapping)
        {
            int cellIndex = 1;
            string cellAddress = string.Empty;
            switch (attribute.MappingOption)
            {
                case MappingOption.FieldName:
                    cellIndex = cellMapping[attribute.MappingKey].GetCellIndex();
                    cellAddress = SpreadsheetExtender.ToCellAddress(rowIndex, cellIndex);
                    break;
                case MappingOption.ColumnName:
                    cellAddress = attribute.MappingKey + rowIndex; break;

                case MappingOption.ColumnIndex:
                    cellAddress = SpreadsheetExtender.ToCellAddress(rowIndex, int.Parse(attribute.MappingKey));
                    break;
            }
            return cellAddress;
        }
        private void SetStyleIndexByCellFormatOption(Cell cell, object value, CellFormatOption option, MappingAttribute attribute)
        {
            int styleIndex = 0;
            switch (cellFormatOption)
            {
                case CellFormatOption.KeepStyle:
                    break;
                case CellFormatOption.OverWrite:
                    styleIndex = SpreadsheetExtender.GetCellStyleIndex(styleSheet, value);
                    cell.SetCellStyle(styleIndex);
                    break;
                case CellFormatOption.ReferenceStyle:
                    if (string.IsNullOrEmpty(attribute.ReferenceStyle))
                    {
                        styleIndex = SpreadsheetExtender.GetCellStyleIndex(styleSheet, value);
                    } else
                    {
                        Cell refCell = SpreadsheetExtender.GetCell(worksheet, attribute.ReferenceStyle);
                        if (refCell == null)
                        {
                            throw new ArgumentException("Please confirm the cell to reference is not null!");
                        }
                        if (refCell.StyleIndex == null || (refCell.StyleIndex != null && !refCell.StyleIndex.HasValue))
                        {
                            styleIndex = SpreadsheetExtender.GetCellStyleIndex(styleSheet, value);
                        } else
                        {
                            styleIndex = int.Parse(refCell.StyleIndex.ToString());
                        }
                        cell.SetCellStyle(styleIndex);
                    } break;
            }
        }
        private void SetStyleIndexByCellFormatOption(Cell cell, object value, CellFormatOption option)
        {
            int styleIndex = 0;
            switch (option)
            {
                case CellFormatOption.KeepStyle:
                    if (cell.StyleIndex != null && cell.StyleIndex.HasValue) return;
                    if ((cell.StyleIndex == null || (cell.StyleIndex != null && !cell.StyleIndex.HasValue)))
                        styleIndex = SpreadsheetExtender.GetCellStyleIndex(styleSheet, value);
                    cell.SetCellStyle(styleIndex); break;
                case CellFormatOption.OverWrite:
                case CellFormatOption.ReferenceStyle:
                    styleIndex = SpreadsheetExtender.GetCellStyleIndex(styleSheet, value);
                    cell.SetCellStyle(styleIndex); break;
            }
        }

        private string GetCellValueWithMapping(int rowIndex, MappingAttribute attribute, Dictionary<string, Cell> cellMapping, SharedStringTablePart sharedTablePart)
        {
            string cellValue = string.Empty;
            switch (attribute.MappingOption)
            {
                case MappingOption.FieldName:
                    Cell cell = cellMapping[attribute.MappingKey];
                    int cellIndex = cell.GetCellIndex();
                    cellValue = worksheet.GetCellValue(rowIndex, cellIndex, sharedTablePart);
                    break;
                case MappingOption.ColumnName:
                    cellValue = worksheet.GetCellValue(attribute.MappingKey + rowIndex.ToString(), sharedTablePart);
                    break;
                case MappingOption.ColumnIndex:
                    cellValue = worksheet.GetCellValue(rowIndex, int.Parse(attribute.MappingKey), sharedTablePart);
                    break;
            }
            return cellValue;
        }
        private object ParserFromString(Type type, string value)
        {
            object target = null;
            switch (type.Name)
            {

                case "Int32":
                    int outValue = 0;
                    Int32.TryParse(value, out outValue);
                    target = (object)outValue;
                    break;
                case "Int64":
                    target = (object)Convert.ToInt64(value);
                    break;


                case "Single":
                    target = (object)Convert.ToSingle(value);
                    break;
                case "Double":
                    target = (object)Convert.ToDouble(value);
                    break;
                case "Decimal":
                    target = (object)Convert.ToDecimal(value);
                    break;
                case "String":

                    target = (object)value; break;
                case "Boolean":
                    target = (object)Convert.ToBoolean(value);
                    break;
                case "DateTime":
                    target = (object)Convert.ToDateTime(value);
                    break;
            }
            return target;
        }


        # endregion
    }




}