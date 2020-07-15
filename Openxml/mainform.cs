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
        public mainform()
        {
            InitializeComponent();
        }

        

        private void mainform_Load(object sender, EventArgs e)
        {
            string filepath = @"D:\word_smp完成.docx";
            //string filepath = @"D:\test.docx";
            //createfile(@"D:\file.docx");
            //addstring(@"D:\file.docx", "hello");
            /*WriteToWordDoc(filepath, "this is a text");
            InsertTableInDoc(filepath);
            var styles = ExtractStylesPart(filepath, true);

            // If the part was retrieved, send the contents to the console.
            if (styles != null)
                Console.WriteLine(styles.ToString());*/
            //getfonts(filepath);
            //FindHeadingParagraphs(filepath);
            XWord xWord = GetDocument(filepath);
            GetParagraph(xWord);
        }

        #region 废物代码
        public static void createfile(string filePath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainDocumentPart = doc.AddMainDocumentPart();
                mainDocumentPart.Document = new Document();
                Body body = mainDocumentPart.Document.AppendChild(new Body());
                Paragraph paragraph = body.AppendChild(new Paragraph());
                Run run = paragraph.AppendChild(new Run());
                run.AppendChild(new Text("this is a new document"));

            }

        }
        public static void addstring(string filePath, string str)
        {

            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                Body body = doc.MainDocumentPart.Document.Body;
                foreach (var paragraph in body.Elements<Paragraph>())
                {
                    Console.WriteLine(paragraph.InnerText);
                }
            }

        }

        public static void WriteToWordDoc(string filepath, string txt)
        {
            // Open a WordprocessingDocument for editing using the filepath.
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true))
            {
                // Assign a reference to the existing document body.
                Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

                // Add a paragraph with some text.
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text(txt));
            }
        }


        public static void InsertTableInDoc(string filepath)
        {
            // Open a WordprocessingDocument for editing using the filepath.
            using (WordprocessingDocument wordprocessingDocument =
                 WordprocessingDocument.Open(filepath, true))
            {
                // Assign a reference to the existing document body.
                Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

                // Create a table.
                Table tbl = new Table();

                // Set the style and width for the table.
                TableProperties tableProp = new TableProperties();
                TableStyle tableStyle = new TableStyle() { Val = "TableGrid" };

                // Make the table width 100% of the page width.
                TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };

                // Apply
                tableProp.Append(tableStyle, tableWidth);
                tbl.AppendChild(tableProp);

                // Add 3 columns to the table.
                TableGrid tg = new TableGrid(new GridColumn(), new GridColumn(), new GridColumn());
                tbl.AppendChild(tg);

                // Create 1 row to the table.
                TableRow tr1 = new TableRow();

                // Add a cell to each column in the row.
                TableCell tc1 = new TableCell(new Paragraph(new Run(new Text("1"))));
                TableCell tc2 = new TableCell(new Paragraph(new Run(new Text("2"))));
                TableCell tc3 = new TableCell(new Paragraph(new Run(new Text("3"))));
                tr1.Append(tc1, tc2, tc3);

                // Add row to the table.
                tbl.AppendChild(tr1);

                // Add the table to the document
                body.AppendChild(tbl);
            }
        }

        // Extract the styles or stylesWithEffects part from a 
        // word processing document as an XDocument instance.
        public static XDocument ExtractStylesPart(
          string fileName,
          bool getStylesWithEffectsPart = true)
        {
            // Declare a variable to hold the XDocument.
            XDocument styles = null;

            // Open the document for read access and get a reference.
            using (var document =
                WordprocessingDocument.Open(fileName, false))
            {
                // Get a reference to the main document part.
                var docPart = document.MainDocumentPart;

                // Assign a reference to the appropriate part to the
                // stylesPart variable.
                StylesPart stylesPart = null;
                if (getStylesWithEffectsPart)
                    stylesPart = docPart.StylesWithEffectsPart;
                else
                    stylesPart = docPart.StyleDefinitionsPart;

                // If the part exists, read it into the XDocument.
                if (stylesPart != null)
                {
                    using (var reader = XmlNodeReader.Create(
                      stylesPart.GetStream(FileMode.Open, FileAccess.Read)))
                    {
                        // Create the XDocument.
                        styles = XDocument.Load(reader);
                    }
                }
            }
            // Return the XDocument instance.
            return styles;
        }




        private RunProperties GetRunPropertyFromParagraph(Paragraph paragraph)
        {
            var runProperties = new RunProperties();
            var fontname = "Calibri";
            var fontSize = "18";
            try
            {
                fontname =
                    paragraph.GetFirstChild<ParagraphProperties>()
                             .GetFirstChild<ParagraphMarkRunProperties>()
                             .GetFirstChild<RunFonts>()
                             .Ascii;
                Console.WriteLine(fontname);
            }
            catch
            {
                //swallow
            }
            try
            {
                fontSize =
                    paragraph.GetFirstChild<Paragraph>()
                             .GetFirstChild<ParagraphProperties>()
                             .GetFirstChild<ParagraphMarkRunProperties>()
                             .GetFirstChild<FontSize>()
                             .Val;
                Console.WriteLine(fontname);
            }
            catch
            {
                //swallow
            }
            runProperties.AppendChild(new RunFonts() { Ascii = fontname });
            runProperties.AppendChild(new FontSize() { Val = fontSize });
            return runProperties;
        }

        private void getfonts(string filename)
        {
            using (WordprocessingDocument wordprocessingDocument =
                    WordprocessingDocument.Open(filename, true))
            {
                //string wordcontent = wordprocessingDocument.MainDocumentPart.Document.Body.InnerText;
                string wordcontent = wordprocessingDocument.MainDocumentPart.Document.Body.InnerText;
                // get all fonts of the word document 
                var fonts = wordprocessingDocument.MainDocumentPart.Document.Descendants<RunFonts>().Select(c => c.Ascii.HasValue ? c.Ascii.InnerText : string.Empty).Distinct().ToList();
            }
        }

        public void FindHeadingParagraphs(string filename)
        {

            var paragraphs = new List<Paragraph>();

            // Open the file read-only since we don't need to change it.
            using (var wordprocessingDocument = WordprocessingDocument.Open(filename, false))
            {
                paragraphs = wordprocessingDocument.MainDocumentPart.Document.Body
                    .OfType<Paragraph>().ToList();
            }
            foreach (Paragraph paragraph in paragraphs)
            {
                Console.WriteLine(paragraph.InnerText);
            }
        }
        #endregion

        /*正文开始*/


        /// <summary>
        /// 获得xml文件
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        /// 
        public XWord GetDocument(string fileName)
        {
            XWord document = new XWord();


            const string documentRelationshipType =
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
            const string stylesRelationshipType =
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
            const string numberingRelationshipType =
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
            const string wordmlNamespace =
              "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace w = wordmlNamespace;

            XDocument xDoc = null;
            XDocument styleDoc = null;
            XDocument numberingDoc = null;

            using (Package wdPackage = Package.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                PackageRelationship docPackageRelationship =
                  wdPackage
                  .GetRelationshipsByType(documentRelationshipType)
                  .FirstOrDefault();
                if (docPackageRelationship != null)
                {
                    Uri documentUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative),docPackageRelationship.TargetUri);//"/word/document.xml"
                    PackagePart documentPart = wdPackage.GetPart(documentUri);
                    
                    //  Load the document XML in the part into an XDocument instance.  
                    xDoc = XDocument.Load(XmlReader.Create(documentPart.GetStream()));

                    //  Find the styles part. There will only be one.  
                    PackageRelationship styleRelation =
                      documentPart.GetRelationshipsByType(stylesRelationshipType)
                      .FirstOrDefault();
                    if (styleRelation != null)
                    {
                        Uri styleUri = PackUriHelper.ResolvePartUri(documentUri, styleRelation.TargetUri);//"/word/styles.xml
                        PackagePart stylePart = wdPackage.GetPart(styleUri);

                        //  Load the style XML in the part into an XDocument instance.  
                        styleDoc = XDocument.Load(XmlReader.Create(stylePart.GetStream()));
                    }
                    //  Find the numbering part. There will only be one.  
                    PackageRelationship numberingRelation =
                      documentPart.GetRelationshipsByType(numberingRelationshipType)
                      .FirstOrDefault();
                    if (numberingRelation != null)
                    {
                        Uri numberingUri = PackUriHelper.ResolvePartUri(documentUri, numberingRelation.TargetUri);//"/word/styles.xml
                        PackagePart numberingPart = wdPackage.GetPart(numberingUri);

                        //  Load the numbering XML in the part into an XDocument instance.  
                        numberingDoc = XDocument.Load(XmlReader.Create(numberingPart.GetStream()));
                    }
                }
            }
            document.xdoc = xDoc;
            document.xname = w;
            document.styledoc = styleDoc;
            document.numberingdoc = numberingDoc;
            return document;
        }


        /// <summary>
        /// 获取段落文字
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        public string ParagraphText(XElement e)
        {
            XNamespace w = e.Name.Namespace;
            //获取段落全部正文
            return e
                   .Elements(w + "r")
                   .Elements(w + "t")
                   .StringConcatenate(element => (string)element);
        }

        /// <summary>
        /// 获取带圈字符和合并字符的段落
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        public string FiledText(XElement e)
        {
            XNamespace w = e.Name.Namespace;

            List<ChineseType> temp = new List<ChineseType>();
            //获取所有insert文字
            string FiledText=e
                   .Elements(w + "r")
                   .Elements(w + "instrText")
                   .StringConcatenate(element => (string)element);

            if ((temp = ChineseList(FiledText)) != null)//如果非空
            {
                string filedtext = "";
                foreach (ChineseType type in temp)
                    filedtext += type.str + "," + type.type +(type.type=="带圈字符"?","+type.symbol:"")+ "|||||";//如果是带圈字符，获取圈的字符类型
                return filedtext;
            }
            else
                return null;

        }

        /// <summary>
        /// 获取拼音指南的段落
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        public string PinYinText(XElement e)
        {
            XNamespace w = e.Name.Namespace;
            //按照rsidR对run进行分类
            var rsidR = e
                    .Elements(w + "r")
                    .ToLookup(x => (string)x.Attributes(w + "rsidR").FirstOrDefault());

            string pinyinText = "";
            
            foreach (var ruby in rsidR)
            {
                pinyinText += ruby.Elements(w + "ruby")
                                  .Elements(w + "rubyBase")
                                  .Elements(w + "r")
                                  .Elements(w + "t")
                                  .StringConcatenate(element => (string)element);
                if (pinyinText == "")//如果没有拼音文字
                    continue;
                pinyinText += ",拼音指南|||||";
            }

            return pinyinText;        
        }


        /// <summary>
        /// 获取段落keyword字体信息
        /// </summary>
        /// <param name="e"></param>
        /// <param name="w"></param>
        /// <param name="keyword"></param>
        /// <returns></returns>
        public FontPr FontText(XElement e,XNamespace w,string keyword)
        {
            string temp="";
            XElement tempe;
            XAttribute tempa;
            //获取所有的run
            var text = from t in e
                       .Elements(w + "r")
                       select t;
            //获取字体属性
            var font = (from f in text
                       where ((temp=(string)f.Element(w + "t"))!=null?temp.Contains(keyword):false)//筛选t正文中包含keyword的run属性
                        select f.Element(w + "rPr")).FirstOrDefault();

            FontPr fontPr = new FontPr();
            fontPr.rFonts = (tempa = GetAttribute(font, "rFonts", "ascii", w)) != null ? (string)tempa : "Times New Roman";//中文字体
            fontPr.bold = (tempe = GetElement(font, "b", w)) != null ? (string)tempe : null;//粗体
            fontPr.italic = (tempe = GetElement(font, "i", w)) != null ? (string)tempe :null;//倾斜
            fontPr.color = (tempa = GetAttribute(font, "color", "val", w)) != null ? (string)tempa :null;//颜色
            fontPr.fontsize = (tempa = GetAttribute(font, "sz", "val", w)) != null ? (double)tempa : 0;//字体大小
            fontPr.underline = (tempa = GetAttribute(font, "u", "val", w)) != null ? (string)tempa : null;//下划线
            fontPr.emphasis = (tempa = GetAttribute(font, "em", "val", w)) != null ? (string)tempa : null;//着重号
            fontPr.vertAlign= (tempa = GetAttribute(font, "vertAlign", "val", w)) != null ? (string)tempa : null;//效果
            fontPr.spacing=(tempa = GetAttribute(font, "spacing", "val", w)) != null ? (double)tempa : 0;//字符间距
            fontPr.postion = (tempa = GetAttribute(font, "postion", "val", w)) != null ? (double)tempa : 0;//字符位置
            fontPr.combine=(tempa=GetAttribute(font, "eastAsianLayout","combine",w))!=null?(string)tempa: null;//双行合一
            fontPr.vert= (tempa = GetAttribute(font, "eastAsianLayout", "vert", w)) != null ? (string)tempa : null;//纵横混排

            return fontPr;
        }


        /// <summary>
        /// 获取元素
        /// </summary>
        /// <param name="e"></param>
        /// <param name="element"></param>
        /// <param name="w"></param>
        /// <returns></returns>
        public XElement GetElement(XElement e,string element,XNamespace w)
        {
            if (e == null)
                return null;
            else
                //返回要求元素
                return e.Element(w + element);  
        }

        /// <summary>
        /// 获取属性
        /// </summary>
        /// <param name="e"></param>
        /// <param name="element"></param>
        /// <param name="attribute"></param>
        /// <param name="w"></param>
        /// <returns></returns>
        public XAttribute GetAttribute(XElement e,string element,string attribute,XNamespace w)
        {
            XElement temp;
            //如果没有要求元素则返回null
            if ((temp=GetElement(e, element, w)) == null)
                return null;
            //返回元素的要求属性
            else
            {
                //如果对应元素没有要求属性
                /*if (temp == null)
                    return null;
                else*/
                    return temp.Attribute(w + attribute);
            }
        }

        /// <summary>
        /// 获取项目等级的属性
        /// </summary>
        /// <param name="numid"></param>
        /// <param name="lvlnum"></param>
        /// <param name="numberingDoc"></param>
        /// <param name="w"></param>
        /// <returns></returns>
        public string GetlvlText(int numid,int lvlnum,XDocument numberingDoc,XNamespace w)
        {
            //如果不存在
            if (numid == -1 || lvlnum == -1)
                return "null";

            //返回numbering.xml里正文中numpr里的numid所对应的abstractnum值
            int abstractnum =
                (int)(from num in numberingDoc.Root.Elements(w + "num")
                where (int)num.Attribute(w + "numId") == numid
                select num.Element(w + "abstractNumId").Attribute(w + "val")).FirstOrDefault();

            //返回numbering.xml里所对应的abstractnum的相关属性
            var abstractn =
                    from abstractNum in numberingDoc.Root.Elements(w + "abstractNum")
                    where (int)abstractNum.Attribute(w + "abstractNumId") == abstractnum
                    select abstractNum;

            //返回numpr里的ilvl的值所对应的lvlText属性
            string lvlText =
                (string)(
                from lvl in abstractn.Descendants(w + "lvl")
                where (int)lvl.Attribute(w + "ilvl") == lvlnum
                select lvl.Element(w + "lvlText").Attribute(w+"val")
                ).FirstOrDefault();

            return lvlText;
        }

        /// <summary>
        /// 获取所有段落
        /// </summary>
        /// <param name="document"></param>
        public void GetParagraph(XWord document)
        {
            string Keyword = "另一个纵横";
            XDocument styleDoc = document.styledoc;
            XNamespace w = document.xname;
            XDocument xDoc = document.xdoc;
            XDocument numberingDoc = document.numberingdoc;
            Element element = new Element();
            
            
            element.init();
            /* string defaultStyle =
                     (string)(
                         from style in styleDoc.Root.Elements(w + "style")
                         where (string)style.Attribute(w + "type") == "paragraph" &&
                               (string)style.Attribute(w + "default") == "1"
                         select style
                     ).First().Attribute(w + "styleId");*/
            //



            double twip = 567.0;//换算量度
            /*A twip (twentieth of a point) is a measure used in laying out space or defining objects on a page 
             * or other area that is to be printed or displayed on a computer screen. 
             * A twip is 1/1440th of an inch or 1/567th of a centimeter. 
             * That is, there are 1440 twips to an inch or 567 twips to a centimeter. 
             * The twip is 1/20th of a point, a traditional measure in printing. 
             * A point is approximately 1/72nd of an inch.*/

            //获取页面属性
            var pages =
                from page in xDoc.Root.Descendants(w + "body")
                let sectpr = page != null ? page.Element(w + "sectPr") : null//节属性

                let pgmar = sectpr != null ? sectpr.Element(w + "pgMar") : null//页边距
                let pgsz = sectpr != null ? sectpr.Element(w + "pgSz") : null//页面大小
                let orient = sectpr != null ? sectpr.Element(w + "orient") : null//方向
                select new
                {
                    pgMar_top = pgmar != null ? ((int)pgmar.Attribute(w + "top")/ twip).ToString("#0.00") : "0",//上边距
                    //pgMar_top = pgmar != null ? (string)pgmar.Attribute(w + "top") : "null",
                    pgMar_bottom = pgmar != null ? ((int)pgmar.Attribute(w + "bottom") / twip).ToString("#0.00") : "0",//下边距
                    pgMar_left = pgmar != null ? ((int)pgmar.Attribute(w + "left") / twip).ToString("#0.00") : "0",//左边距
                    pgMar_right = pgmar != null ? ((int)pgmar.Attribute(w + "right") / twip).ToString("#0.00") : "0",//右边距
                    pgSz_width = pgsz != null ? Math.Round((int)pgsz.Attribute(w + "w") / twip,2): 0,//页宽
                    pgSz_height = pgsz != null ? Math.Round((int)pgsz.Attribute(w + "h") / twip,2) : 0,//页长
                    orient = orient != null ? (string)orient.Attribute(w + "val") : "Portrait"//方向
                };

            //获取段落属性
            var paras =
                from para in xDoc
                             .Root
                             .Element(w + "body")
                             .Descendants(w + "p")
                let ppr = para.Element(w + "pPr")
                //------------------------------------------------------------------
                //页面设置
                //------------------------------------------------------------------

                let sectpr = ppr != null ? ppr.Element(w + "sectPr") : null//节属性
                let pgmar = sectpr != null ? sectpr.Element(w + "pgMar") : null//页边距
                let pgsz = sectpr != null ? sectpr.Element(w + "pgSz") : null//纸张大小
                let orient = sectpr != null ? sectpr.Element(w + "orient") : null//方向

                //------------------------------------------------------------------
                //段落排版
                //------------------------------------------------------------------

                //-----------------
                //常规
                //-----------------                
                let jc = ppr != null ? ppr.Element(w + "jc") : null//对齐方式
             
                let outlineLvl = ppr != null ? ppr.Element(w + "outlineLvl") : null//大纲等级

                //-----------------
                //缩进
                //-----------------
                let ind = ppr != null ? ppr.Element(w + "ind") : null//缩进
                let ind_left = ind != null ? ind.Attribute(w + "left") : null//左缩进
                let ind_leftChars = ind != null ? ind.Attribute(w + "leftChars") : null//左缩进（字符）
                let ind_right = ind != null ? ind.Attribute(w + "right") : null//右缩进
                let ind_rightChars = ind != null ? ind.Attribute(w + "rightChars") : null//右缩进（字符）               
                let ind_firstline = ind != null ? ind.Attribute(w + "firstLine") : null//首行缩进
                let ind_firstlineChars = ind != null ? ind.Attribute(w + "firstLineChars") : null//首行缩进（字符）
                let ind_hanging = ind != null ? ind.Attribute(w + "hanging") : null//悬挂缩进
                let ind_hangingChars = ind != null ? ind.Attribute(w + "hangingChars") : null//悬挂缩进（字符）

                //-----------------
                //间距
                //-----------------
                let spacing = ppr != null ? ppr.Element(w + "spacing") : null//间距
                let spacing_beforeLines = spacing != null ? spacing.Attribute(w + "beforeLines") : null//段前
                let spacing_afterLines = spacing != null ? spacing.Attribute(w + "afterLines") : null//段后
                let spacing_line = spacing != null ? spacing.Attribute(w + "line") : null//行距

                //------------------------------------------------------------------
                //段落格式
                //------------------------------------------------------------------

                //-----------------
                //分栏
                //-----------------
                let cols = sectpr != null ? sectpr.Element(w + "cols") : null//分栏
                let cols_num = cols != null ? cols.Attribute(w + "num") : null//栏数
                let cols_sep = cols != null ? cols.Attribute(w + "sep") : null//分隔线


                //let pbdr = ppr != null ? ppr.Element(w + "pBdr") : null  //边框
                //let split = pbdr != null ? ppr.Element(w + "bottom") : null  //下框

                //-----------------
                //首字下沉
                //-----------------
                let framepr = ppr != null ? ppr.Element(w + "framePr") : null//首字下沉
                let framepr_dropCap = framepr != null ? framepr.Attribute(w + "dropCap") : null//位置
                let framepr_lines = framepr != null ? framepr.Attribute(w + "lines") : null//行数

                //-----------------
                //项目符号
                //-----------------
                let numPr = ppr != null ? ppr.Element(w + "numPr") : null//项目符号
                let numPr_ilvl = numPr != null ? numPr.Element(w + "ilvl") : null//等级
                let numPr_numId = numPr != null ? numPr.Element(w + "numId") : null//索引值


                select new
                {
                    //------------------------------------------------------------------
                    //页面设置
                    //------------------------------------------------------------------

                    //-----------------
                    //页边距
                    //-----------------
                    pgMar_top = pgmar != null ? (string)pgmar.Attribute(w + "top") : null,//上边距
                    pgMar_bottom = pgmar != null ? (string)pgmar.Attribute(w + "bottom") : null,//下边距
                    pgMar_left = pgmar != null ? (string)pgmar.Attribute(w + "left") : null,//左边距
                    pgMar_right = pgmar != null ? (string)pgmar.Attribute(w + "right") : null,//右边距

                    //-----------------
                    //纸张大小
                    //-----------------
                    pgSz_width = pgsz != null ? (string)pgsz.Attribute(w + "w") : null,//页宽
                    pgSz_height = pgsz != null ? (string)pgsz.Attribute(w + "h") : null,//页长

                    //-----------------
                    //方向
                    //-----------------
                    orient = orient != null ? (string)orient.Attribute(w + "val") : "Portrait",//方向

                    //------------------------------------------------------------------
                    //段落排版
                    //------------------------------------------------------------------

                    //-----------------
                    //常规
                    //-----------------
                    jc = jc != null ? (string)jc.Attribute(w + "val") : "both",//对齐方式
                    outlineLvl = outlineLvl != null ? (int)outlineLvl.Attribute(w + "val") : -1,//样式

                    //-----------------
                    //缩进
                    //-----------------
                    ind_left = ind_left != null ? (double)ind_left : 0,//左缩进
                    ind_leftChars = ind_leftChars != null ? (double)ind_leftChars : 0,//左缩进（字符）
                    ind_right = ind_right != null ? (double)ind_right : 0,//右缩进
                    ind_rightChars = ind_rightChars != null ? (double)ind_rightChars : 0,//右缩进（字符）
                    ind_firstline = ind_firstline != null ? (double)ind_firstline : 0,//首行缩进
                    ind_firstlineChars = ind_firstlineChars != null ? (double)ind_firstlineChars : 0,//首行缩进（字符）
                    ind_hanging = ind_hanging != null ? (double)ind_hanging : 0,//悬挂缩进
                    ind_hangingChars = ind_hangingChars != null ? (double)ind_hangingChars : 0,//悬挂缩进（字符）

                    //-----------------
                    //间距
                    //-----------------
                    spacing_beforeLines = spacing_beforeLines != null ? (double)spacing_beforeLines : 0,//段前
                    spacing_afterLines = spacing_afterLines != null ? (double)spacing_afterLines : 0,//段后
                    spacing_line = spacing != null ? (double)spacing_line : 0,//行距

                    //------------------------------------------------------------------
                    //段落格式
                    //------------------------------------------------------------------

                    //-----------------
                    //分栏(目前没有偏左和偏右)
                    //-----------------
                    cols_num = cols_num != null ? (int)cols_num :1,//栏数
                    cols_sep = cols_sep != null ? (int)cols_sep :0 ,//分隔线

                    //-----------------
                    //首字下沉
                    //-----------------
                    framepr_dropCap = framepr_dropCap != null ? (string)framepr_dropCap : null,//位置
                    framepr_lines = framepr_lines != null ? (int)framepr_lines : 0,//行数
                    //-----------------
                    //项目符号(目前无法识别图案问题，只能识别等级)
                    //-----------------
                    numPr_ilvl = numPr != null ? (int)numPr_ilvl.Attribute(w + "val") : -1,//等级
                    numPr_numId = numPr != null ? (int)numPr_numId.Attribute(w + "val") : -1,//索引值

                    //------------------------------------------------------------------
                    //字符排版
                    //------------------------------------------------------------------

                    //-----------------
                    //字体-字符间距-中文版式(双行和一,纵横混排)
                    //-----------------
                    Fonts = FontText(para, w, Keyword),//中文字体、字形、字号、颜色、下划线、效果、着重号、间距、位置、双行合一、纵横混排
                    //------------------------------------------------------------------
                    //中文版式
                    //------------------------------------------------------------------

                    //-----------------
                    //带圈字符和合并字符
                    //-----------------
                    FiledText = FiledText(para),//带圈字符、合并字符

                    //-----------------
                    //拼音指南
                    //-----------------
                    PinYinText = PinYinText(para),//拼音指南

                    //------------------------------------------------------------------
                    //段落正文
                    //------------------------------------------------------------------
                    ParagraphNode = para,//段落编号
                    Text = ParagraphText(para),//段落正文

                };
        
                #region 废物
            //纸张大小
            // Find all paragraphs in the document.  
            /*  var paragraphs =
                  from para in xDoc
                               .Root
                               .Element(w + "body")
                               .Descendants(w + "p")
                  let border_top = para
                                 .Elements(w + "pPr")
                                 .Elements(w + "pBdr")
                                 .Elements(w + "top")
                                 .FirstOrDefault()
                  select new
                  {
                      ParagraphNode = para,
                      Border_line = border_top != null ? (string)border_top.Attribute(w + "val") : defaultStyle
                      //element.border.val= border != null ? (string)border.Attribute(w + "val") : defaultStyle
                  };*/

            // Retrieve the text of each paragraph.  
            /* var paraWithText =
                 from para in paragraphs
                 select new
                 {
                     ParagraphNode = para.ParagraphNode,
                     Text = ParagraphText(para.ParagraphNode),
                     Border_line=para.Border_line
                 };
             RichTextBox richTextBox = new RichTextBox();*/


            /*var paraPage =
                from para in paragraphs
                select new
                {
                    ParagraphNode = para.ParagraphNode,
                    Text = ParagraphText(para.ParagraphNode),
                };


            foreach (var p in paraWithText)
                Console.WriteLine("StyleName: >{0}< Border_line:{1}",  p.Text,p.Border_line);
            foreach (var p in paras)
                Console.WriteLine("pgmar:{0},{1},{2},{3},pgsz:{4},{5},orient:{6},{7}", p.pgMar_top, p.pgMar_bottom, p.pgMar_left, p.pgMar_right, p.pgSz_height,
                    p.pgSz_width, p.orient, p.Text);*/
                #endregion



            string result = "";
            ///页面属性
            foreach (var p in pages) {
                result = string.Format("页边距\n上：{0}cm 下：{1}cm 左：{2}cm 右：{3}cm;\n", p.pgMar_top, p.pgMar_bottom, p.pgMar_left, p.pgMar_right);
                result += string.Format("方向和纸张大小\n方向：{0} 纸张大小 宽：{1}cm 高 {2}cm {3}\n\n", OrientType(p.orient),p.pgSz_width, p.pgSz_height,PageSizeType(p.pgSz_width, p.pgSz_height));
            }
            
            //段落属性
            foreach (var p in paras)
            {
               
                result += p.Text+"\n";
                result += string.Format("段落排版\n常规\n对齐方式：{0} 样式：{1}\n", JsType(p.jc), outlineLvlType(p.outlineLvl));
                result += string.Format("缩进\n左：{0} 右：{1} 特殊格式：{2}\n",indCount((int)UnitType.ch, p.ind_left,p.ind_leftChars),
                    indCount((int)UnitType.ch, p.ind_right, p.ind_rightChars),ind_special(p.ind_firstline+p.ind_firstlineChars,p.ind_hanging+p.ind_hangingChars) );
                result += string.Format("间距\n段前：{0}行 段后：{1}行 行距：{2}倍行距\n\n", UnitTypeChanged((int)UnitType.ch, p.spacing_beforeLines),
                    UnitTypeChanged((int)UnitType.ch, p.spacing_afterLines), UnitTypeChanged((int)UnitType.line, p.spacing_line)) ;
                result += string.Format("段落格式\n分栏\n栏数：{0} 分隔线：{1}\n", colsNum(p.cols_num),colsSep(p.cols_sep));
                result += string.Format("首字下沉\n位置：{0} 行数：{1}行\n", frameprDropType(p.framepr_dropCap), p.framepr_lines);
                result += string.Format("项目符号\n符号：{0},{1}\n\n", numPrilvl(p.numPr_ilvl), GetlvlText(p.numPr_numId,p.numPr_ilvl,numberingDoc,w));
                result += string.Format("字符排版\n字体\n中文字体：{0} 字形：{1} 字号：{2}磅 颜色：{3} 下划线：{4} 效果：{5} 着重号：{6}\n",
                    p.Fonts.rFonts,boldOritalic(p.Fonts.bold,p.Fonts.italic),p.Fonts.fontsize/2, 
                    hexTorgb(p.Fonts.color),underlineType(p.Fonts.underline),vertAlignType(p.Fonts.vertAlign),emphasisType(p.Fonts.emphasis));
                result += string.Format("字符间距\n间距：{0}{1}磅 位置：{2}{3}磅\n", spacingType(p.Fonts.spacing), Math.Abs(p.Fonts.spacing / 20),
                    postionType(p.Fonts.postion), Math.Abs(p.Fonts.postion / 2));
                
                result += string.Format("中文版式\n{0}{1}{2}{3}\n\n", p.FiledText,p.PinYinText,combineType(p.Fonts.combine),vertType(p.Fonts.vert));
            }
            rtxt.Text = result;
            

            //项目符号显示
            //comboBox1.Font = new System.Drawing.Font("Wingdings", comboBox1.Font.Size);
            //comboBox1.Text = Regex.Unescape("\u006E");
            
        }

    }
}
