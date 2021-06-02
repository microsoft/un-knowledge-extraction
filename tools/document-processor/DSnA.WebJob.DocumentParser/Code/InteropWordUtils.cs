//Copyright(c) Microsoft Corporation.All rights reserved.
//Licensed under the MIT License.

using Microsoft.Office.Interop.Word;
using System;

namespace DSnA.WebJob.DocumentParser
{
    public interface IInteropWordUtils
    {
        Application CreateWordAppInstance();
        Document OpenDocument(string file, Application wordApp);
        void DisposeIneropObject(Application wordApp, bool saveChanges = false);
    }

    class InteropWordUtils : IInteropWordUtils
    {
        /// <summary>
        /// Creates word application instance
        /// </summary>
        /// <returns></returns>
        public Application CreateWordAppInstance()
        {
            return new Application
            {
                DisplayAlerts = WdAlertLevel.wdAlertsNone,
                Visible = false,
                Options = { SavePropertiesPrompt = false, SaveNormalPrompt = false, DisplayPasteOptions = false, DoNotPromptForConvert = true }
            };
        }

        /// <summary>
        /// Opens word document
        /// </summary>
        /// <param name="file"></param>
        /// <param name="wordApp"></param>
        /// <returns></returns>
        public Document OpenDocument(string file, Application wordApp)
        {
            return wordApp.Documents.Open(file, ReadOnly: false);
        }

        /// <summary>
        /// Disposes all COM Objects
        /// </summary>
        /// <param name="wordApp"></param>
        public void DisposeIneropObject(Application wordApp,bool saveChanges = false)
        {
            try
            {
                wordApp.Quit(SaveChanges: saveChanges);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception)
            {

                wordApp.Quit(SaveChanges: false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}
