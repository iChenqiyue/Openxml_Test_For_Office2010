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
        public class XWord
        {

            public XNamespace xname;
            public XDocument xdoc;
            public XDocument styledoc;
        }

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
            const string wordmlNamespace =
              "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace w = wordmlNamespace;

            XDocument xDoc = null;
            XDocument styleDoc = null;

            using (Package wdPackage = Package.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                PackageRelationship docPackageRelationship =
                  wdPackage
                  .GetRelationshipsByType(documentRelationshipType)
                  .FirstOrDefault();
                if (docPackageRelationship != null)
                {
                    Uri documentUri =
                        PackUriHelper
                        .ResolvePartUri(
                           new Uri("/", UriKind.Relative),
                                 docPackageRelationship.TargetUri);
                    PackagePart documentPart =
                        wdPackage.GetPart(documentUri);

                    //  Load the document XML in the part into an XDocument instance.  
                    xDoc = XDocument.Load(XmlReader.Create(documentPart.GetStream()));

                    //  Find the styles part. There will only be one.  
                    PackageRelationship styleRelation =
                      documentPart.GetRelationshipsByType(stylesRelationshipType)
                      .FirstOrDefault();
                    if (styleRelation != null)
                    {
                        Uri styleUri = PackUriHelper.ResolvePartUri(documentUri, styleRelation.TargetUri);
                        PackagePart stylePart = wdPackage.GetPart(styleUri);

                        //  Load the style XML in the part into an XDocument instance.  
                        styleDoc = XDocument.Load(XmlReader.Create(stylePart.GetStream()));
                    }
                }
            }
            document.xdoc = xDoc;
            document.xname = w;
            document.styledoc = styleDoc;
            return document;
        }





        public static string ParagraphText(XElement e)
        {
            XNamespace w = e.Name.Namespace;
            return e
                   .Elements(w + "r")
                   .Elements(w + "t")
                   .StringConcatenate(element => (string)element);
        }





        public void GetParagraph(XWord document)
        {
            XDocument styleDoc = document.styledoc;
            XNamespace w = document.xname;
            XDocument xDoc = document.xdoc;
            Element element = new Element();
            Common common = new Common();
            
            
            element.init();
            /* string defaultStyle =
                     (string)(
                         from style in styleDoc.Root.Elements(w + "style")
                         where (string)style.Attribute(w + "type") == "paragraph" &&
                               (string)style.Attribute(w + "default") == "1"
                         select style
                     ).First().Attribute(w + "styleId");*/
            //


            double twip = 567.0;
            /*A twip (twentieth of a point) is a measure used in laying out space or defining objects on a page 
             * or other area that is to be printed or displayed on a computer screen. 
             * A twip is 1/1440th of an inch or 1/567th of a centimeter. 
             * That is, there are 1440 twips to an inch or 567 twips to a centimeter. 
             * The twip is 1/20th of a point, a traditional measure in printing. 
             * A point is approximately 1/72nd of an inch.*/
            var pages =
                from page in xDoc.Root.Descendants(w + "body")
                let sectpr = page != null ? page.Element(w + "sectPr") : null
                let pgmar = sectpr != null ? sectpr.Element(w + "pgMar") : null
                let pgsz = sectpr != null ? sectpr.Element(w + "pgSz") : null
                let orient = sectpr != null ? sectpr.Element(w + "orient") : null
                select new
                {
                    pgMar_top = pgmar != null ? ((int)pgmar.Attribute(w + "top")/ twip).ToString("#0.00") : "0",
                    //pgMar_top = pgmar != null ? (string)pgmar.Attribute(w + "top") : "null",
                    pgMar_bottom = pgmar != null ? ((int)pgmar.Attribute(w + "bottom") / twip).ToString("#0.00") : "0",
                    pgMar_left = pgmar != null ? ((int)pgmar.Attribute(w + "left") / twip).ToString("#0.00") : "0",
                    pgMar_right = pgmar != null ? ((int)pgmar.Attribute(w + "right") / twip).ToString("#0.00") : "0",
                    pgSz_width = pgsz != null ? Math.Round((int)pgsz.Attribute(w + "w") / twip,2): 0,
                    pgSz_height = pgsz != null ? Math.Round((int)pgsz.Attribute(w + "h") / twip,2) : 0,
                    orient = orient != null ? (string)orient.Attribute(w + "val") : "Portrait"
                };

            foreach (var p in pages)
                Console.WriteLine("pgmar:{0},{1},{2},{3},pgsz:{4},{5},orient:{6}", p.pgMar_top, p.pgMar_bottom, p.pgMar_left, p.pgMar_right, p.pgSz_height,
                    p.pgSz_width, p.orient);


            var paras =
                from para in xDoc
                             .Root
                             .Element(w + "body")
                             .Descendants(w + "p")
                let ppr = para.Element(w + "pPr")
                //页面设置
                let sectpr = ppr != null ? ppr.Element(w + "sectPr") : null
                let pgmar = sectpr != null ? sectpr.Element(w + "pgMar") : null  //页边距
                let pgsz = sectpr != null ? sectpr.Element(w + "pgSz") : null  //纸张大小
                let orient = sectpr != null ? sectpr.Element(w + "orient") : null  //方向
                //段落排版
                let jc = ppr != null ? ppr.Element(w + "jc") : null  //常规
                let outlineLvl = ppr != null ? ppr.Element(w + "outlineLvl") : null
                let ind = ppr != null ? ppr.Element(w + "ind") : null  //缩进
                let ind_left = ind != null ? ind.Attribute(w + "left") : null  //缩进
                let ind_right = ind != null ? ind.Attribute(w + "right") : null  //缩进
                let ind_leftChars = ind != null ? ind.Attribute(w + "leftChars") : null  //缩进
                let ind_rightChars = ind != null ? ind.Attribute(w + "rightChars") : null  //缩进
                let ind_firstline = ind != null ? ind.Attribute(w + "firstLine") : null  //缩进
                let ind_hanging = ind != null ? ind.Attribute(w + "hanging") : null  //缩进
                let ind_firstlineChars = ind != null ? ind.Attribute(w + "firstLineChars") : null  //缩进
                let ind_hangingChars = ind != null ? ind.Attribute(w + "hangingChars") : null  //缩进


                let spacing = ppr != null ? ppr.Element(w + "spacing") : null //间距
                let spacing_beforeLines=spacing != null ? spacing.Attribute(w + "beforeLines") : null //间距
                let spacing_afterLines = spacing != null ? spacing.Attribute(w + "afterLines") : null //间距
                let spacing_line = spacing != null ? spacing.Attribute(w + "line") : null //间距


                //段落格式
                let cols = sectpr != null ? sectpr.Element(w + "cols") : null  //分栏
                let pbdr = ppr != null ? ppr.Element(w + "pBdr") : null  //边框
                let split = pbdr != null ? ppr.Element(w + "bottom") : null  //下框
                let framepr = ppr != null ? ppr.Element(w + "framePr") : null//下沉
                let numPr = ppr != null ? ppr.Element(w + "numPr") : null//项目符号
                let numPr_ilvl = numPr != null ? numPr.Element(w + "ilvl") : null//项目符号
                let numPr_numId = numPr != null ? numPr.Element(w + "numId") : null//项目符号


                select new
                {
                    //页面设置
                    //页边距
                    pgMar_top = pgmar != null ? (string)pgmar.Attribute(w + "top") : "null",
                    pgMar_bottom = pgmar != null ? (string)pgmar.Attribute(w + "bottom") : "null",
                    pgMar_left = pgmar != null ? (string)pgmar.Attribute(w + "left") : "null",
                    pgMar_right = pgmar != null ? (string)pgmar.Attribute(w + "right") : "null",
                    //纸张大小
                    pgSz_width = pgsz != null ? (string)pgsz.Attribute(w + "w") : "null",
                    pgSz_height = pgsz != null ? (string)pgsz.Attribute(w + "h") : "null",
                    
                    //方向
                    orient = orient != null ? (string)orient.Attribute(w + "val") : "Portrait",

                    
                    //段落排版
                    //常规
                    jc = jc != null ? (string)jc.Attribute(w + "val") : "both",
                    outlineLvl = outlineLvl != null ? (int)outlineLvl.Attribute(w + "val") : -1,
                    //缩进
                    ind_left = ind_left != null ? (double)ind_left : 0,
                    ind_leftChars = ind_leftChars != null ? (double)ind_leftChars : 0,
                    ind_right = ind_right != null ? (double)ind_right : 0,
                    ind_rightChars = ind_rightChars != null ? (double)ind_rightChars : 0,
                    ind_firstline = ind_firstline != null ? (double)ind_firstline : 0,
                    ind_firstlineChars = ind_firstlineChars != null ? (double)ind_firstlineChars : 0,
                    ind_hanging = ind_hanging != null ? (double)ind_hanging : 0,
                    ind_hangingChars = ind_hangingChars != null ? (double)ind_hangingChars : 0,
                    spacing_beforeLines = spacing_beforeLines != null ? (double)spacing_beforeLines : 0,
                    spacing_afterLines = spacing_afterLines != null ? (double)spacing_afterLines : 0,
                    spacing_line = spacing != null ? (double)spacing_line : 0,


                    //段落格式
                    //分栏
                    cols = cols != null ? (string)cols.Attribute(w + "num") : "0",
                    split = split != null ? (string)split.Attribute(w + "val") : "0",
                    //首字下沉
                    framepr_dropCap = framepr != null ? (string)framepr.Attribute(w + "dropCap") : "0",
                    framepr_lines = framepr != null ? (string)framepr.Attribute(w + "lines") : "0",
                    //项目符号
                    numPr_ilvl = numPr != null ? (string)numPr_ilvl.Attribute(w + "val") : "0",
                    numPr_numId = numPr != null ? (string)numPr_numId.Attribute(w + "val") : "0",

                    ParagraphNode = para,
                    Text = ParagraphText(para)

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
            foreach (var p in pages) {
                result = string.Format("页边距\n上：{0}cm 下：{1}cm 左：{2}cm 右：{3}cm;\n", p.pgMar_top, p.pgMar_bottom, p.pgMar_left, p.pgMar_right);
                result += string.Format("方向和纸张大小\n方向：{0} 纸张大小 宽：{1}cm 高 {2}cm {3}\n", common.OrientType(p.orient),p.pgSz_width, p.pgSz_height,common.PageSizeType(p.pgSz_width, p.pgSz_height));
            }
            foreach (var p in paras)
            {
                result += p.Text+"\n";
                result += string.Format("段落排版\n常规\n对齐方式：{0} 样式：{1}\n", common.JsType(p.jc), common.outlineLvlType(p.outlineLvl));
                result += string.Format("缩进\n左：{0} 右：{1} 特殊格式：{2}\n",common.indCount(0,p.ind_left,p.ind_leftChars),
                    common.indCount(0, p.ind_right, p.ind_rightChars),common.ind_special(p.ind_firstline+p.ind_firstlineChars,p.ind_hanging+p.ind_hangingChars) );
                result += string.Format("间距\n段前：{0}行 段后：{1}行 行距：{2}倍行距\n", common.UnitTypeChanged(0,p.spacing_beforeLines),
                    common.UnitTypeChanged(0, p.spacing_afterLines), common.UnitTypeChanged(5, p.spacing_line));
                result += string.Format("段落格式\n分栏\n栏数：{0} 分割线：{1}\n", p.cols,p.split);
                result += string.Format("首字下沉\n位置：{0} 行数：{1}\n", p.framepr_dropCap, p.framepr_lines);
                result += string.Format("项目符号\n符号：{0},{1}\n\n", p.numPr_ilvl, p.numPr_numId);

            }
            rtxt.Text = result;
        }

    }
}
