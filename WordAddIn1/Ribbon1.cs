using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Text.RegularExpressions;

// TODO:   按照以下步骤启用功能区(XML)项: 

// 1. 将以下代码块复制到 ThisAddin、ThisWorkbook 或 ThisDocument 类中。

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. 在此类的“功能区回调”区域中创建回调方法，以处理用户
//    操作(如单击某个按钮)。注意: 如果已经从功能区设计器中导出此功能区，
//    则将事件处理程序中的代码移动到回调方法并修改该代码以用于
//    功能区扩展性(RibbonX)编程模型。

// 3. 向功能区 XML 文件中的控制标记分配特性，以标识代码中的相应回调方法。  

// 有关详细信息，请参见 Visual Studio Tools for Office 帮助中的功能区 XML 文档。


namespace WordAddIn1
{

    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        Word.Application wordApp;
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
        }

        #region IRibbonExtensibility 成员

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("WordAddIn1.Ribbon1.xml");
        }

        #endregion

        #region 功能区回调
        //在此处创建回叫方法。有关添加回叫方法的详细信息，请访问 https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void jiben(Office.IRibbonControl control)
        {
            wordApp = Globals.ThisAddIn.Application;
            Word.Paragraph par = wordApp.Application.ActiveDocument.Paragraphs[1];
            par.Range.Font.NameFarEast = "方正小标宋_GBK";
            par.Range.Font.Size = 22;
            par.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            par.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
            par.Range.ParagraphFormat.LineSpacing = 30;
        }

        public void wTitle(Office.IRibbonControl control)
        {
            wordApp = Globals.ThisAddIn.Application;
            Word.Paragraphs par=wordApp.Application.ActiveDocument.Paragraphs;
            foreach(Word.Paragraph p in par)
            {
                p.Range.Font.NameFarEast = "方正仿宋_GBK";
                p.Range.Font.NameAscii = "宋体";
                p.Range.Font.Size = 16;
                string pattern1 = "[一二三四五六七八九十]{1,3}、";
                string pattern2 = "^[（(][一二三四五六七八九十]{1,3}[）)]";
                string pattern3 = "[一二三四五六七八九十]{1,3}、.*。";
                string pattern4 = "^[（(][一二三四五六七八九十]{1,3}[）)].*。";

                if (System.Text.RegularExpressions.Regex.IsMatch(p.Range.Text, pattern3))
                {
                    p.Range.Sentences[1].Font.NameFarEast= "方正黑体_GBK";
                }
                else if (System.Text.RegularExpressions.Regex.IsMatch(p.Range.Text, pattern4))
                {
                    p.Range.Sentences[1].Font.NameFarEast = "方正楷体_GBK";
                } else if(p.Range.Text.Length < 40 && System.Text.RegularExpressions.Regex.IsMatch(p.Range.Text, pattern1)){
                    p.Range.Font.NameFarEast = "方正黑体_GBK";
                }else if (p.Range.Text.Length < 40 && System.Text.RegularExpressions.Regex.IsMatch(p.Range.Text, pattern2))
                {
                    p.Range.Font.NameFarEast = "方正楷体_GBK";
                }
                p.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
            }
        }

        public void DelPar(Office.IRibbonControl control)
        {
            wordApp = Globals.ThisAddIn.Application;
            Word.Paragraphs par = wordApp.Application.ActiveDocument.Paragraphs;
            foreach(Word.Paragraph p in par)
            {
                if (p.Range.Text.Length == 1)
                {
                    p.Range.Delete();
                }
            }
        }
        public void ZBName(Office.IRibbonControl control)
        {
            wordApp = Globals.ThisAddIn.Application;
            Word.Paragraphs par = wordApp.Application.ActiveDocument.Paragraphs;
            par[par.Count].Range.InsertAfter("\n\n\n");
            wordApp.Application.ActiveDocument.Paragraphs[wordApp.Application.ActiveDocument.Paragraphs.Count].Range.InsertAfter("隆阳区住房和城乡建设局住房保障中心\n");
            wordApp.Application.ActiveDocument.Paragraphs[wordApp.Application.ActiveDocument.Paragraphs.Count].Range.InsertAfter(DateTime.Now.Year+"年"+ DateTime.Now.Month + "月"+DateTime.Now.Day + "日");
            Word.Paragraphs par2= wordApp.Application.ActiveDocument.Paragraphs;
            par2[par2.Count].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            par2[par2.Count-1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
        }

        public void ZJName(Office.IRibbonControl control)
        {
            wordApp = Globals.ThisAddIn.Application;
            Word.Paragraphs par = wordApp.Application.ActiveDocument.Paragraphs;
            par[par.Count].Range.InsertAfter("\n\n\n");
            wordApp.Application.ActiveDocument.Paragraphs[wordApp.Application.ActiveDocument.Paragraphs.Count].Range.InsertAfter("隆阳区住房和城乡建设局\n");
            wordApp.Application.ActiveDocument.Paragraphs[wordApp.Application.ActiveDocument.Paragraphs.Count].Range.InsertAfter(DateTime.Now.Year + "年" + DateTime.Now.Month + "月" + DateTime.Now.Day + "日");
            Word.Paragraphs par2 = wordApp.Application.ActiveDocument.Paragraphs;
            par2[par2.Count].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            par2[par2.Count - 1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
        }


        #endregion

        #region 帮助器

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
