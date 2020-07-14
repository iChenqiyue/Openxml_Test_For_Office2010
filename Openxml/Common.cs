using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Openxml
{
    public partial class mainform : Form
    {


        string[] Outlinelvl = { "一", "二", "三", "四", "五", "六", "七", "八", "九" };
        string[] col = { "一栏", "两栏", "三栏", "偏左", "偏右" };
        enum UnitType { ch, cm, mm, inch, pt, line };
        string[] Unit = { "字符", "厘米", "毫米", "英寸", "磅", "行" };

        #region 页面大小
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
        PageSize A3 = new PageSize("A3", 29.7, 42);
        PageSize A4 = new PageSize("A4", 21, 29.7);
        PageSize A5 = new PageSize("A5", 14.8, 21);

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
        #endregion

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

        /// <summary>
        /// 判断段落对齐方式
        /// </summary>
        /// <param name="js"></param>
        /// <returns></returns>
        public string JsType(string js)
        {
            if (js == "left")
                return "左对齐";
            else if (js == "right")
                return "右对齐";
            else if (js == "center")
                return "居中";
            else if (js == "both")
                return "两端对齐";
            else if (js == "distribute")
                return "分散对齐";
            else
                return js;
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
        public class FontPr
        {
            public string rFonts;
            public string bold;
            public string italic;
            public string color;
            public double fontsize;
            public string underline;
            public string vertAlign;
            public string emphasis;
            public double spacing;
            public double postion;
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
        public class ChineseType
        {
            public string type;
            public string str;
            public string symbol;
        }

        // eq \o\ac(○,圈) eq \o\ac(□,A) eq \o\ac(△,!) eq \o\ac(◇,壹)
        //eq \o(\s\up 8(合并),\s\do 3(字符))eq \o(\s\up 8(两　),\s\do 3(　字))eq \o(\s\up 8(三个),\s\do 3(字))eq \o(\s\up 8(五个),\s\do 3(合并字))eq \o(\s\up 8(六个合),\s\do 3(并字符))

        public List<ChineseType> ChineseList(string chars)
        {

            if (chars == null)
                return null;
            else
            {
                List<ChineseType> list = new List<ChineseType>();

                string temp = "";

                if (chars.Contains(@" eq \o\ac("))
                {
                    temp = chars.Replace(@" eq \o\ac(", "");

                    temp = temp.Replace(")", ";");
                    string[] strs = temp.Split(';');
                    foreach (string str in strs)
                    {
                        if (str == "")
                            break;
                        ChineseType chinessType = new ChineseType();
                        chinessType.type = "带圈字符";
                        chinessType.str = str.Substring(2);
                        chinessType.symbol = str.Substring(0, 1);
                        
                        list.Add(chinessType);
                    }

                    return list;
                }
                else if (chars.Contains(@"eq \o(\s\up 8("))
                {
                    temp = chars.Replace(@"eq \o(\s\up 8(", "");
                    temp = temp.Replace(@"),\s\do 3(", "");
                    temp = temp.Replace("　　", "");
                    temp = temp.Replace("))", ";");
                    string[] strs = temp.Split(';');
                    foreach (string str in strs)
                    {
                        ChineseType chinessType = new ChineseType();
                        chinessType.type = "合并字符";
                        chinessType.str = str;
                        list.Add(chinessType);
                    }
                    return list;
                }
                else
                    return null;
            }          
        }
        #endregion
    }
}
