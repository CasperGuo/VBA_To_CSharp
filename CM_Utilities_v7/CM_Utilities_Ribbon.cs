using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new CM_Utilities_Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace CM_Utilities_v7
{
    [ComVisible(true)]
    public class CM_Utilities_Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public CM_Utilities_Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("CM_Utilities_v7.CM_Utilities_Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void Clean_Up_Riders_Ribbon(Office.IRibbonControl rbnCtrl)
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Document currentDoc = Globals.ThisAddIn.Application.ActiveDocument;

            int intRidersTotal=0;

            const string strRIDER_NAME_TOKEN = "-";
            const string strRIDER_HEADER = "Schedule ot College Board";

            /*try
            {
            */
                if (!currentDoc.TrackRevisions)
                    currentDoc.TrackRevisions = true;

                if(currentDoc.Fields.Count!=0)
                {
                    foreach (Word.Field fld in currentDoc.Fields)
                    {
                        /* Each Rider is a separate paragraph with a field at the end
                         * The purpose of the code is to determine if that field is expanded
                         * to display the ENTIRE RIDER or if it is just a field at the end of
                         * a lone paragraph.  If lone paragraph, then delete.
                        */
                        Word.Paragraph para = fld.Result.Paragraphs[1];
                        if(fld.Type==Word.WdFieldType.wdFieldIf)
                        {
                            /* Have to look at paragraph range in order to get to
                             * the characters selection. I could have stuck .range at the end of Paragraphs(1)
                             * above, however that line of code is already 3 dots deep (a reference of
                             * a reference of a reference) and that's bad coding practice
                             */
                            Word.Range rngFld = para.Range;
                            rngFld.Select();
                            /* Each rider name starts with a dash (-) and is highlighted
                             * Added 02/11/2016, after testing against old riders, the OR clause.
                             * May delete in the coming months when no old riders
                             */
                             if((rngFld.Characters[1].Text==strRIDER_NAME_TOKEN
                                && rngFld.Characters[1].HighlightColorIndex==Word.WdColorIndex.wdBrightGreen
                                && app.Selection.Paragraphs.Count==1)
                                ||
                                (rngFld.Characters[1].HighlightColorIndex==Word.WdColorIndex.wdBrightGreen
                                && app.Selection.Paragraphs.Count==1))
                                {
                                    app.Selection.Paragraphs[1].Range.Delete();
                                    intRidersTotal++;
                                }
                        }
                        else
                        {
                            /* Get rid of the highlighted Rider Names here
                             * During testing on 02/11/2016 noticed this ALONE works to clean up riders
                             */
                            app.Selection.Fields[1].Unlink();
                            app.Selection.Paragraphs[1].Range.Delete();
                            app.Selection.Find.Execute(strRIDER_HEADER);
                            if (app.Selection.Find.Found)
                                app.Selection.ParagraphFormat.PageBreakBefore = 1;
                        }
                    }

                    if (intRidersTotal == 0)
                    {
                        MessageBox.Show("No Riders exist in this document:\n");
                    }
                    else
                        MessageBox.Show("Number of Unnecessary Riders Found: " + intRidersTotal);
                }
            //}

            /*catch (Exception)
            {
                throw;
            }
            */
        }
        public void MakeHEDAmendment_Ribbon(Office.IRibbonControl rbnCtrl) { }
        public void CreateSoleSourceLetter_Ribbon(Office.IRibbonControl rbnCtrl) { }
        public void RefreshShortcuts_Ribbon(Office.IRibbonControl rbnCtrl) { }
        public void DeleteMyRoad_Ribbon(Office.IRibbonControl rbnCtrl) { }
        public void FormatPrice_Ribbon(Office.IRibbonControl rbnCtrl) { }
        public void FormatDateSpellOutMonth_Ribbon(Office.IRibbonControl rbnCtrl) { }
        public void FormatPhoneNumber_Ribbon(Office.IRibbonControl rbnCtrl) { }
        public void InterfaceForSpellNumber_Ribbon(Office.IRibbonControl rbnCtrl) { }
        public bool GetEnabled(Office.IRibbonControl rbnCtrl)
        {
            return true;
        }

        #endregion

        #region Helpers

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
