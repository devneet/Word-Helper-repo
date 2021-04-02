using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;


namespace Word_Helper
{
    public class WordHelperClass
    {
        
        public String insertPicture(string wordDocPath, string imageFilePath)
        {

            try
            {

                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document wordDoc = wordApp.Documents.Open(wordDocPath);
                wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                wordApp.Visible = false;

                wordDoc.InlineShapes.AddPicture(imageFilePath);

                wordDoc.Save();
                wordApp.Quit();

                wordDoc = null;
                wordApp = null;

                return "SUCCESS : Activity got executed.";

            }
            catch (Exception ex)
            {

                return "EXCEPTION : "+ex.InnerException.Message.ToString();

            }

        }

    }
}

