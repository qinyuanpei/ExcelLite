using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

namespace ExcelLite.Format
{
    public class CellFormatBuilder
    {
        /// <summary>
        /// 存储CellFormat成员的字典
        /// </summary>
        private Dictionary<Styles, OpenXmlElement> formatDict;
        /// <summary>
        /// 默认Excel单元格样式
        /// </summary>
        private CellFormat defaultCellFormat;
        /// <summary>
        /// Excel样式表
        /// </summary>
        private Stylesheet styleSheet;
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="styleSheet"></param>
        public CellFormatBuilder(Stylesheet styleSheet)
        {
            this.styleSheet = styleSheet;
            this.formatDict = new Dictionary<Styles, OpenXmlElement>();
            this.defaultCellFormat = GetDefaultCellFormat(styleSheet); 
        }
        /// <summary>
        ///  获取默认单元格式
        /// </summary>
        /// <param name="styleSheet"></param>
        /// <returns></returns>
        public static CellFormat GetDefaultCellFormat(Stylesheet styleSheet)
        {
            if(styleSheet==null)
                throw new ArithmeticException("StyleSheet does not exist in this Excel document!");
            var cellFormats=styleSheet.CellFormats;
            if(cellFormats.Count<=0)
            {
               throw new ArithmeticException("CellFormat does not exist in this Excel StyleSheet!");
            
            }
        
                  return cellFormats.Elements<CellFormat>().First();
        }
       /// <summary>
       /// 获取默认数值样式
       /// </summary>
       /// <param name="styleSheet"></param>
       /// <param name="styleIndex"></param>
       /// <returns></returns>
        public static CellFormat GetDefaultNumberFormat(Stylesheet styleSheet, out int styleIndex)
        {

            Dictionary<StyleParams, string> values = new Dictionary<StyleParams, string>();
            values.Add(StyleParams.NumberFormat, "General");
            OpenXmlElement style = CellStyleFactory.Create(Styles.Number, values);
            //生成单元格格式
            CellFormatBuilder builder = new CellFormatBuilder(styleSheet);
            builder.AppendStyle(Styles.Number, style);
            return builder.Build(out styleIndex); 
        }
        /// <summary>
        /// 获取默认日期样式
        /// </summary>
        /// <param name="styleSheet"></param>
        /// <param name="styleIndex"></param>
        /// <returns></returns>
        public static CellFormat GetDefaultDateFormat(Stylesheet styleSheet, out int styleIndex)
        {
            Dictionary<StyleParams, string> values = new Dictionary<StyleParams, string>();
            values.Add(StyleParams.NumberFormat, "yyyy/m/d");
            OpenXmlElement style = CellStyleFactory.Create(Styles.Number,values);
            //生成单元格格式
            CellFormatBuilder builder = new CellFormatBuilder(styleSheet);
            builder.AppendStyle(Styles.Number, style);
            return builder.Build(out styleIndex);
        
        }
        /// <summary>
        /// 添加Excel样式，相同键名的样式按照添加顺序覆盖
        /// </summary>
        /// <param name="styleType"></param>
        /// <param name="style"></param>
        public void AppendStyle(Styles styleType, OpenXmlElement style)
        {
            if (!formatDict.ContainsKey(styleType))
                this.formatDict.Add(styleType, style);
            this.formatDict[styleType] = style;
        
        }
        public CellFormat Build(out int cellStyleIndex)
        {
            Fill fill = null;
            UInt32Value fillID;
            Font font = null;
            UInt32Value fontID;
            Border border = null;
            UInt32Value borderID;
            Alignment align = null;
            UInt32Value alignID;
            NumberingFormat nf = null;
            UInt32Value nfID;
            if(!formatDict.ContainsKey(Styles.Fill)){
                fillID=defaultCellFormat.FillId; 
            }
            else
            {
               fill=(Fill)formatDict[Styles.Fill];
               styleSheet.Fills.Append(fill);
               fillID=new UInt32Value((uint)QueryElementID(Styles.Fill,fill)); 
            }

            if (!formatDict.ContainsKey(Styles.Font)) {
                fontID = defaultCellFormat.FillId;            
            }
            else{
                font = (Font)formatDict[Styles.Font];
                styleSheet.Fonts.Append(font);
                fontID =new UInt32Value((uint)QueryElementID(Styles.Font,font)); 
            
            }
              
            if (!formatDict.ContainsKey(Styles.Border)) {
                borderID = defaultCellFormat.BorderId;            
            }
            else{
                border = (Border)formatDict[Styles.Border];
                styleSheet.Borders.Append(border);
                borderID =new UInt32Value((uint)QueryElementID(Styles.Border,border));             
            }

            if (!formatDict.ContainsKey(Styles.Align))
            {
                alignID = defaultCellFormat.BorderId;
            }
            else
            {
                align = (Alignment)formatDict[Styles.Align];
            }

            if (!formatDict.ContainsKey(Styles.Number))
            {
                nfID = defaultCellFormat.NumberFormatId;
            }
            else
            {
                if (styleSheet.NumberingFormats == null)
                    styleSheet.NumberingFormats = new NumberingFormats();
                nf = (NumberingFormat)formatDict[Styles.Number];
                nf.NumberFormatId = UInt32Value.FromUInt32((uint)styleSheet.NumberingFormats.Count());
                styleSheet.NumberingFormats.AppendChild(nf);
                nfID = new UInt32Value((uint)QueryElementID(Styles.Number,nf));
            }
            //CellFormat
            CellFormat cellFormat = new CellFormat() { Alignment = align };
            cellFormat.FillId = fillID;
            cellFormat.FontId = fontID;
            cellFormat.BorderId = borderID;
            cellFormat.NumberFormatId = borderID;
            cellFormat.ApplyBorder = true;
            cellFormat.ApplyFill = true;
            cellFormat.ApplyFont = true;
            cellFormat.ApplyNumberFormat = true;
            cellFormat.FormatId = UInt32Value.FromUInt32((uint)styleSheet.CellFormats.Count());
            cellFormat.ApplyNumberFormat = true;

            //save new cellFormat
            styleSheet.CellFormats.Append(cellFormat);
            styleSheet.Save();

            //Return inde of CellFormat
            cellStyleIndex = styleSheet.CellFormats.ToList().IndexOf(cellFormat);
            return cellFormat;

        }
        private int QueryElementID(Styles type, OpenXmlElement element)
        {

            int reval = -1;
            switch (type)
            { 
                case Styles.Border:
                    reval = styleSheet.Fills.ToList().IndexOf(element);
                    break;
                case Styles.Fill:
                    reval = styleSheet.Fills.ToList().IndexOf(element);
                    break;


                case Styles.Font:
                    reval = styleSheet.Fonts.ToList().IndexOf(element);
                    break;
                case Styles.Number:
                    reval = styleSheet.NumberingFormats.ToList().IndexOf(element);
                    break; 
            }
            return reval;
        }
    }
}
