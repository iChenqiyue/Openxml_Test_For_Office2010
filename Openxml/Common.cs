using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace Openxml
{
    public class Common
    {

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
        static PageSize A3 = new PageSize("A3", 29.7, 42);
        static PageSize A4 = new PageSize("A4", 21, 29.7);
        static PageSize A5 = new PageSize("A5", 14.8, 21);

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
            string[] outlinelvl = { "一", "二", "三", "四", "五", "六", "七", "八", "九" };
            for (int i = 0; i < 9; i++)
            {
                if (val == i)
                    return "标题"+outlinelvl[i];
            }
            if (val == -1)
                return "正文";
            else
                return val.ToString();
        }

        enum UnitType { ch,cm,mm,inch,pt,line};
        string[] Unit = { "字符", "厘米", "毫米", "英寸", "磅" ,"行"};
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

        public string ind_special(double firstline,double hanging)
        {
            if (firstline != 0)
                return "首行缩进";
            else if (hanging != 0)
                return "悬挂缩进";
            else
                return "无";
        }

        public string indCount(int type,double init,double chars)
        {
            if (chars != 0)
                return UnitTypeChanged(type, chars).ToString() + Unit[(int)UnitType.ch];
            else if (init != 0)
                return UnitTypeChanged(type, init).ToString() + Unit[type];
            else
                return "0" + Unit[type];
        }

    }
}
