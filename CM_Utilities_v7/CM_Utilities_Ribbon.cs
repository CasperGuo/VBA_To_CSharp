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
            Word.Paragraph fldPara;
            Word.Range fldRider;

            int intRidersTotal=0;
            int intUnnecessaryRiders = 0;
            int intNecessaryRiders = 0;

            // const string strRIDER_NAME_TOKEN = "-";
            const string strRIDER_HEADER = "Schedule to College Board";

            try
            {
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
                        if(fld.Type==Word.WdFieldType.wdFieldIf)
                        {
                            /* This is a much "simplier" and straight forward deletion of the unncessary riders
                             * This is MORE related to the architecture i.e. Merge Field is either "True" Or "False"
                             * So if the field code is "False = True", then delete it, that field
                             * Tested an it works.
                             * Look at old VBA code to see how complicated I made the selection.
                             */
                            fldPara = fld.Result.Paragraphs[1];
                            fldRider = fld.Result;
                            if (fld.Code.Text.Contains("\"False\" = \"True"))
                                {
                                    intUnnecessaryRiders++;
                                }
                            else
                            { 
                                fld.Unlink();
                                intNecessaryRiders++;
                                fldRider.Find.Execute(strRIDER_HEADER);
                                if (fldRider.Find.Found)
                                    fldRider.ParagraphFormat.PageBreakBefore = -1;
                            }
                            fldPara.Range.Select();
                            fldPara.Range.Delete();
                            intRidersTotal++;
                        }
                        else
                        {
                            /* Get rid of the highlighted Rider Names here
                             * During testing on 02/11/2016 noticed this ALONE works to clean up riders
                             */
                        }
                    }

                    if (intRidersTotal == 0)
                    {
                        MessageBox.Show("No Riders exist in this document:\n");
                    }
                    else
                        MessageBox.Show("Number of Unnecessary Riders Found: " + intUnnecessaryRiders + "\n"
                            + "Number of Necessary Riders Found " + intNecessaryRiders + "\n"
                            + "Number of Total Riders Found " + intRidersTotal);
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public void CreateSoleSourceLetter_Ribbon(Office.IRibbonControl rbnCtrl)
        {
            // I think I'm going to try and do the "Web Version of the Sole Source Letter for this one.
        }
        public void MakeHEDAmendment_Ribbon(Office.IRibbonControl rbnCtrl) { }
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
