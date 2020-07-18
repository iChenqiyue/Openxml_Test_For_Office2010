using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using System.Xml.Linq;
using System.IO.Packaging;
using System.IO;

using System.Xml;
using System.Text.RegularExpressions;
using System.Collections;

namespace Openxml
{
    public partial class mainform : Form
    {


        string[] Outlinelvl = { "一", "二", "三", "四", "五", "六", "七", "八", "九" };//项目等级
        string[] col = { "一栏", "两栏", "三栏", "偏左", "偏右" };//分栏种类
        enum UnitType { ch, cm, mm, inch, pt, line };//单位种类
        string[] Unit = { "字符", "厘米", "毫米", "英寸", "磅", "行" };//代为种类对应的中文

        enum LineType { single,dotted,dashSmallGap,dashed,dotDash,dotDotDash,doubleline,triple,thinThickSmallGap,thickThinSmallGap};
        string[] Line = { "single", "dotted", "dashSmallGap", "dashed", "dotDash", "dotDotDash", @"double", "triple", "thinThickSmallGap", "thickThinSmallGap" };

        /// <summary>
        /// xml文档类
        /// </summary>
        public class XWord
        {

            public XNamespace xname;
            public XNamespace xmarkup;
            public XNamespace xworddrawing;
            public XNamespace xdrawing;
            public XNamespace xshape;
            public XDocument xdoc;
            public XDocument styledoc;
            public XDocument numberingdoc;
          
            
        }
        public class ArtText
        {
            public FontPr font;
            public string style;
            public string shape;
            public string surrounding;
            public ArtText(FontPr tfont,string tsurrounding,string tshape)
            {
                font = tfont;
                surrounding = tsurrounding;
                shape = tshape;
            }
        }

        public class Picture
        {
            public double width;
            public double height;
            public string surrounding;
            public Picture() { }
            public Picture(double twidth,double theight,string tsurrounding)
            {
                width = twidth;
                height = theight;
                surrounding = tsurrounding;
            }
        }

        public class Border
        {
            public string shadow;
            public string frame;
            public string linetype;
            public string color;
            public double width;
            public string type;
        }

        #region 页面大小
        /// <summary>
        /// 页面大小类
        /// </summary>
        class PageSize
        {
            public string name;
            public double width;
            public double length;
            public PageSize(string name1, double width1, double length1)
            {
                name = name1;
                width = width1;
                length = length1;
            }
        }

        PageSize A3 = new PageSize("A3", 29.7, 42);//A3纸张
        PageSize A4 = new PageSize("A4", 21, 29.7);//A4纸张
        PageSize A5 = new PageSize("A5", 14.8, 21);//A5纸张

        /// <summary>
        /// 判断页大小类型
        /// </summary>
        /// <param name="width1"></param>
        /// <param name="length1"></param>
        /// <returns></returns>
        public string PageSizeType(double width1, double length1)
        {
            if (width1 == A3.width && length1 == A3.length)
                return A3.name;
            else if (width1 == A4.width && length1 == A4.length)
                return A4.name;
            else if (width1 == A5.width && length1 == A5.length)
                return A5.name;
            else
                return "error";
        }


        /// <summary>
        /// 判断页面方向
        /// </summary>
        /// <param name="orient"></param>
        /// <returns></returns>
        public string OrientType(string orient)
        {
            if (orient == "Portrait")
                return "纵向";
            else if (orient == "Landscape")
                return "横向";
            else
                return orient;
        }
        #endregion



        /// <summary>
        /// 判断段落对齐方式
        /// </summary>
        /// <param name="js"></param>
        /// <returns></returns>
        public string JcType(string jc)
        {
            if (jc == "left")
                return "左对齐";
            else if (jc == "right")
                return "右对齐";
            else if (jc == "center")
                return "居中";
            else if (jc == "both")
                return "两端对齐";
            else if (jc == "distribute")
                return "分散对齐";
            else
                return jc;
        }
        
        /// <summary>
        /// 判断大纲级别
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        public string outlineLvlType(int val)
        {
            
            for (int i = 0; i < 9; i++)
            {
                if (val == i)
                    return "标题"+ Outlinelvl[i];
            }
            if (val == -1)
                return "正文";
            else
                return val.ToString();
        }

        /// <summary>
        /// 换算单位
        /// </summary>
        /// <param name="type"></param>
        /// <param name="d"></param>
        /// <returns></returns>
        public double UnitTypeChanged(int type,double d)
        {
            if (type == (int)UnitType.ch)
                return Math.Round(d / 100,2);
            else if (type == (int)UnitType.cm)
                return Math.Round(d / 567, 2);
            else if (type == (int)UnitType.pt)
                return Math.Round(d / 20, 2);
            else if (type == (int)UnitType.line)
                return Math.Round(d / 240, 2);
            else
                return d;
        }
        
        /// <summary>
        /// 判断特殊缩进
        /// </summary>
        /// <param name="firstline"></param>
        /// <param name="hanging"></param>
        /// <returns></returns>
        public string ind_special(double firstline,double hanging)
        {
            if (firstline != 0)
                return "首行缩进";
            else if (hanging != 0)
                return "悬挂缩进";
            else
                return "无";
        }

        /// <summary>
        /// 计算缩进，两种单位，字符和非字符
        /// </summary>
        /// <param name="type"></param>
        /// <param name="init"></param>
        /// <param name="chars"></param>
        /// <returns></returns>
        public string indCount(int type,double init,double chars)
        {
            if (chars != 0)
                return UnitTypeChanged(type, chars).ToString() + Unit[(int)UnitType.ch];
            else if (init != 0)
                return UnitTypeChanged(type, init).ToString() + Unit[type];
            else
                return "0" + Unit[type];
        }

        /// <summary>
        /// 计算分栏个数
        /// </summary>
        /// <param name="num"></param>
        /// <returns></returns>
        public string colsNum(int num)
        {
            for(int i=1;i<=3; i++)
            {
                if (num == i)
                    return col[i-1];
            }
            return num.ToString();
        }

        /// <summary>
        /// 判断是否有分隔线
        /// </summary>
        /// <param name="sep"></param>
        /// <returns></returns>
        public string colsSep(int sep)
        {
            if (sep == 0)
                return "无";
            else
                return "有";
        }

        /// <summary>
        /// 判断首字下沉类型
        /// </summary>
        /// <param name="dropCap"></param>
        /// <returns></returns>
        public string frameprDropType(string dropCap)
        {
            if (dropCap ==null)
                return "无";
            else if (dropCap == "drop")
                return "下沉";
            else if (dropCap == "margin")
                return "悬挂";
            else
                return dropCap;

        }

        /// <summary>
        /// 判断项目符号级别
        /// </summary>
        /// <param name="ilvl"></param>
        /// <returns></returns>
        public string numPrilvl(int ilvl)
        {
            for (int i = 0; i < 9; i++)
            {
                if (ilvl == i)
                    return Outlinelvl[i]+"级";
            }
            if (ilvl == -1)
                return "无";
            else
                return ilvl.ToString();
        }


        #region 字体相关
        /// <summary>
        /// 字体类
        /// </summary>
        public class FontPr
        {
            public string rFonts;//中文字体
            public string bold;//粗体
            public string italic;//斜体
            public string color;//颜色
            public double fontsize;//字体大小
            public string underline;//下划线
            public string vertAlign;//效果
            public string emphasis;//着重号
            public double spacing;//间距
            public double postion;//位置
            public string combine;//双行合一
            public string vert;//纵横混排
            public Border border;
            public FontPr() { }
            public FontPr(string rfonts,double size)
            {
                rFonts = rfonts;
                fontsize = size;
            }
        }


        

        /// <summary>
        /// 判断字形
        /// </summary>
        /// <param name="bold"></param>
        /// <param name="italic"></param>
        /// <returns></returns>
        public string boldOritalic(string bold,string italic)
        {
            string result="";
            if (bold == "")
                result += "加粗";
            if (italic == "")
                result += "倾斜";
            if (result == "")
                result = "常规";
            return result;

        }

        /// <summary>
        /// 判断下划线类型
        /// </summary>
        /// <param name="underline"></param>
        /// <returns></returns>
        public string underlineType(string underline)
        {
            if (underline == null)
                return "无";
            else if (underline == "single")
                return "单直线";
            else if (underline == "double")
                return "双直线";
            else if (underline == "thick")
                return "粗直线";
            else if (underline == "dotted")
                return "虚线";
            else if (underline == "dottedHeavy")
                return "粗虚线";
            else
                return underline;

        }

        /// <summary>
        /// 判断字体效果
        /// </summary>
        /// <param name="vertAlign"></param>
        /// <returns></returns>
        public string vertAlignType(string vertAlign)
        {
            if (vertAlign == null)
                return "无";
            else if (vertAlign == "superscript")
                return "上标";
            else if (vertAlign == "subscript")
                return "下标";
            else
                return vertAlign;
        }

        /// <summary>
        /// 判断是否有着重号
        /// </summary>
        /// <param name="emphasis"></param>
        /// <returns></returns>
        public string emphasisType(string emphasis)
        {
            if (emphasis == null)
                return "无";
            else
                return "有";
        }

        /// <summary>
        /// 颜色十六进制转RGB
        /// </summary>
        /// <param name="hex"></param>
        /// <returns></returns>
        public string hexTorgb(string hex)
        {
            if (hex == null)
                return null;
            else if (hex == "auto")
                return "auto";
            else
            {
                int r = Convert.ToInt32(hex.Substring(0, 2), 16);
                int g = Convert.ToInt32(hex.Substring(2, 2), 16);
                int b = Convert.ToInt32(hex.Substring(4, 2), 16);
                return string.Format("RGB({0},{1},{2})", r, g, b);
            }
        }

        /// <summary>
        /// 判断间距类型
        /// </summary>
        /// <param name="spacing"></param>
        /// <returns></returns>
        public string spacingType(double spacing)
        {
            if (spacing < 0)
                return "紧缩";
            else if (spacing > 0)
                return "加宽";
            else
                return "标准";
                
        }

        /// <summary>
        /// 判断位置类型
        /// </summary>
        /// <param name="postion"></param>
        /// <returns></returns>
        public string postionType(double postion)
        {
            if (postion < 0)
                return "降低";
            else if (postion > 0)
                return "提升";
            else
                return "标准";

        }

        /// <summary>
        /// 判断是否双行合一
        /// </summary>
        /// <param name="combine"></param>
        /// <returns></returns>
        public string combineType(string combine)
        {
            if (combine != null)
                return "双行合一";
            else
                return "";
        }

        /// <summary>
        /// 判断是否纵横混排
        /// </summary>
        /// <param name="vert"></param>
        /// <returns></returns>
        public string vertType(string vert)
        {
            if (vert != null)
                return "纵横混排";
            else
                return "";
        }

        /// <summary>
        /// 中文版式带圈字符及合并字符类
        /// </summary>
        public class ChineseType
        {
            public string type;//种类
            public string str;//字符串
            public string symbol;//符号
        }

        // eq \o\ac(○,圈) eq \o\ac(□,A) eq \o\ac(△,!) eq \o\ac(◇,壹)
        //eq \o(\s\up 8(合并),\s\do 3(字符))eq \o(\s\up 8(两　),\s\do 3(　字))eq \o(\s\up 8(三个),\s\do 3(字))eq \o(\s\up 8(五个),\s\do 3(合并字))eq \o(\s\up 8(六个合),\s\do 3(并字符))

        /// <summary>
        /// 返回中文字符类型及值列表
        /// </summary>
        /// <param name="chars"></param>
        /// <returns></returns>
        public List<ChineseType> ChineseList(string chars)
        {

            if (chars == null)
                return null;
            else
            {
                List<ChineseType> list = new List<ChineseType>();

                string temp = "";

                if (chars.Contains(@" eq \o\ac("))//包含带圈字符样式
                {
                    temp = chars.Replace(@" eq \o\ac(", "");//替换

                    temp = temp.Replace(")", ";");//替换，以分割
                    string[] strs = temp.Split(';');//分割字符串
                    foreach (string str in strs)
                    {
                        if (str == "")
                            break;
                        ChineseType chinessType = new ChineseType();
                        chinessType.type = "带圈字符";
                        chinessType.str = str.Substring(2);
                        chinessType.symbol = str.Substring(0, 1);
                        
                        list.Add(chinessType);//添加带圈字符到列表
                    }

                    return list;
                }
                else if (chars.Contains(@"eq \o(\s\up 8("))//包含合并字符样式
                {
                    temp = chars.Replace(@"eq \o(\s\up 8(", "");//替换
                    temp = temp.Replace(@"),\s\do 3(", "");
                    temp = temp.Replace("　　", "");
                    temp = temp.Replace("))", ";");
                    string[] strs = temp.Split(';');//分割字符串
                    foreach (string str in strs)
                    {
                        ChineseType chinessType = new ChineseType();
                        chinessType.type = "合并字符";
                        chinessType.str = str;
                        list.Add(chinessType);//添加合并字符到列表
                    }
                    return list;
                }
                else
                    return null;
            }          
        }
        #endregion

        /// <summary>
        /// 边框类型
        /// </summary>
        /// <param name="linetype"></param>
        /// <param name="shadow"></param>
        /// <param name="frame"></param>
        /// <returns></returns>
        public string borderType(Border border)
        {
            if (border.type == "all")
                return "全部";
            else if (border.type == "part")
                return "虚框";
            else if (border.shadow != null)
                return "阴影";
            else if (border.frame != null)
                return "三维";
            else if (border.linetype != null || border.type == "none")
                return "方框";
            else
                return "无";
        }

        


        public int lineType(string linetype)
        {
            for(int i = 0; i < Line.Length; i++)
            {
                if (linetype == Line[i])
                    return i+1;
            }
            return 0;
        }

        public Border TableBorderType(List<Border>borders)
        {

            Border border = new Border();
            if (borders[1].linetype == "none")
            {
                return borders[0];
            }
            else if (borders[0].linetype == borders[1].linetype)
            {
                border = borders[0];
                border.type = "all";
                return border;
            }
            else
            {
                border = borders[0];
                border.type = "part";
                return border;
            }

        }

        
    }
}
