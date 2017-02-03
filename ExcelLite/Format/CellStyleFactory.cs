//文件名：CellStyleFactory.CS
//作者：PayneQin
//日期：2016-08-11
//描述：Excel单元格样式类
//
//备注：
//1.从工厂创建Fill类型样式，需要指定FillColor字段，该字段对应的数值为RGB颜色数值
//2.从工厂创建Font类型样式，需要制定FontSize、FontName、FontFamil、FontColor字段
//3.从工厂创建Align类型样式，需要指定HAlignType、VAlignType字段
//4.从工厂创建Border类型样式，需要指定BorderColor字段
//5.从工厂创建Number类型样式，需要指定NumberFormat字段using System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

namespace ExcelLite.Format
{
    public static class CellStyleFactory
    {
        public static OpenXmlElement Create(Styles type, Dictionary<StyleParams, string> paramsDict)
        {
            if (paramsDict == null || paramsDict.Count == 0)
            {
                throw new ArgumentException("Please check whether there are parameters in the paramsDict");
            }
            OpenXmlElement element = null;
            switch (type)
            {
                case Styles.Fill:
                    if (!paramsDict.ContainsKey(StyleParams.FillColor))
                        throw new ArithmeticException("Argument list does not eist key:FillColor");
                    element = CreateFill(paramsDict[StyleParams.FillColor]); break;
                case Styles.Font:
                    if (!paramsDict.ContainsKey(StyleParams.FontName) ||
                        !paramsDict.ContainsKey(StyleParams.FontColor)
                        || !paramsDict.ContainsKey(StyleParams.FontSize))
                        throw new ArithmeticException("Argument list must  eist key:FontName、FontColor、FontSize");
                    element = CreateFont(paramsDict[StyleParams.FontName], paramsDict[StyleParams.FontSize], paramsDict[StyleParams.FontColor]); break;
                case Styles.Align:
                    if (!paramsDict.ContainsKey(StyleParams.HAlignType) ||
                        !paramsDict.ContainsKey(StyleParams.VAlignType))
                        throw new ArithmeticException("Argument list must  eist key:FontName、FontColor、FontSize");
                    element = CreateAlign(paramsDict[StyleParams.HAlignType], paramsDict[StyleParams.VAlignType]); break;

                case Styles.Border:
                    if (!paramsDict.ContainsKey(StyleParams.BorderColor))
                        throw new ArithmeticException("Argument list must  eist key:BorderColor");
                    element = CreateBorder(paramsDict[StyleParams.BorderColor]); break;
                case Styles.Number:
                    if (!paramsDict.ContainsKey(StyleParams.NumberFormat))
                        throw new ArithmeticException("Argument list must  eist key:NumberFormat");
                    element = CreateNumberingFormat(paramsDict[StyleParams.NumberFormat]); break;
            }
            return element;
        }
        /// <summary>
        /// 创建一个字体相关的Excel样式
        /// </summary>
        /// <param name="fontName"></param>
        /// <param name="fontSize"></param>
        /// <param name="fontColor"></param>
        /// <returns></returns>
        private static Font CreateFont(string fontName, string fontSize, string fontColor)
        {
            Font font = new Font();
            font.FontName = new FontName() { Val = new StringValue(fontName) };
            font.FontSize = new FontSize() { Val = new DoubleValue(double.Parse(fontSize)) };
            font.Color = new Color() { Rgb = HexBinaryValue.FromString(fontColor) };
            font.FontFamilyNumbering = new FontFamilyNumbering { Val = 2 };
            return font;
        }
        private static Fill CreateFill(string color)
        {
            PatternFill patternFill = new PatternFill() { PatternType = PatternValues.Solid };
            patternFill.Append(new ForegroundColor() { Rgb = HexBinaryValue.FromString(color) });
            patternFill.Append(new BackgroundColor() { Indexed = (UInt32Value)64U });
            Fill fill = new Fill() { PatternFill = patternFill };
            return fill;
        }

        private static Alignment CreateAlign(string alignH, string alignV)
        {
            Alignment align = new Alignment();
            if (align.Vertical == null)
                align.Vertical = new EnumValue<VerticalAlignmentValues>();
            if (align.Horizontal == null)
                align.Horizontal = new EnumValue<HorizontalAlignmentValues>();
            switch (alignH)
            {
                case "Left":
                    align.Horizontal.Value = HorizontalAlignmentValues.Left;
                    break;
                case "Right":
                    align.Horizontal.Value = HorizontalAlignmentValues.Right;
                    break;
                case "Center":
                    align.Horizontal.Value = HorizontalAlignmentValues.Center;
                    break;
            }
            switch (alignV)
            {
                case "Top":
                    align.Vertical.Value = VerticalAlignmentValues.Top;
                    break;
                case "Bottom":
                    align.Vertical.Value = VerticalAlignmentValues.Bottom;
                    break;
                case "Center":
                    align.Vertical.Value = VerticalAlignmentValues.Center;
                    break;
            }
            return align;
        }

        private static Border CreateBorder(string color)
        {
            Border border = new Border();
            border.LeftBorder = new LeftBorder()
            {

                Color = new Color() { Rgb = HexBinaryValue.FromString(color) },
                Style = BorderStyleValues.Thin
            };
            border.RightBorder = new RightBorder()
            {

                Color = new Color() { Rgb = HexBinaryValue.FromString(color) },
                Style = BorderStyleValues.Thin
            };
            border.BottomBorder = new BottomBorder()
            {

                Color = new Color() { Rgb = HexBinaryValue.FromString(color) },
                Style = BorderStyleValues.Thin
            };

            return border;
        }
        private static NumberingFormat CreateNumberingFormat(string formatCode)
        {
            NumberingFormat nf = new NumberingFormat();
            nf.FormatCode = StringValue.FromString(formatCode);
            return nf;
        }
    }

    public enum Styles
    {
        /// <summary>
        /// 填充单元格
        /// </summary>
        Fill,
        /// <summary>
        /// 单元格字体
        /// </summary>
        Font,
        /// <summary>
        /// 单元格对齐
        /// </summary>
        Align,
        /// <summary>
        /// 单元格边框
        /// </summary>
        Border,
        /// <summary>
        /// 数字样式
        /// </summary>
        Number,
    }
    public enum StyleParams
    {
        /// <summary>
        /// 填充颜色
        /// </summary>
        FillColor,
        /// <summary>
        /// 字体名称
        /// </summary>            
        FontName,
        /// <summary>
        /// 字体大小
        /// </summary>
        FontSize,
        /// <summary>
        /// 水平对齐方式
        /// </summary>
        HAlignType,
        /// <summary>
        /// 垂直对齐方式
        /// </summary>
        VAlignType,
        /// <summary>
        /// 边框颜色
        /// </summary>
        BorderColor,
        /// <summary>
        /// zityse
        /// </summary>
        FontColor,
        /// <summary>
        /// 数字格式
        /// </summary>
        NumberFormat
    }

}
