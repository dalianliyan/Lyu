using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SGJ.Utilities
{
    public class WordReporter
    {

        #region Microsoft Word

        public static string WriteWord(string pstrFilePath, Dictionary<string, string> pdicTitleBlock)
        {
            if (!string.IsNullOrWhiteSpace(pstrFilePath) && File.Exists(pstrFilePath) && pdicTitleBlock != null)
            {
                WordprocessingDocument lobjApp = WordprocessingDocument.Open(pstrFilePath, true);
                MainDocumentPart lobjDocPart = lobjApp.MainDocumentPart;
                Dictionary<string, SdtElement> ldicWordControl = new Dictionary<string, SdtElement>();
                // get all named control in word
                foreach (SdtElement sdt in lobjDocPart.Document.Descendants<SdtElement>())
                {
                    SdtAlias alias1 = sdt.Descendants<SdtAlias>().FirstOrDefault();
                    if ((alias1 != null))
                    {
                        string sdtTitle = alias1.Val.Value;
                        ldicWordControl.Add(sdtTitle, sdt);
                    }
                }

                // write each title block in dictionary
                foreach (KeyValuePair<string, string> lobjTitleBlock in pdicTitleBlock)
                {
                    if (ldicWordControl.ContainsKey(lobjTitleBlock.Key))
                    {
                        WriteControl(ldicWordControl[lobjTitleBlock.Key], lobjTitleBlock.Value);
                    }
                }
                lobjApp.Close();
            }
            return pstrFilePath;
        }

        public static bool WriteControl(SdtElement lobjSdtElement, string lstrText)
        {

            if (lobjSdtElement is SdtBlock)
            {
                //skip
            }
            else if (lobjSdtElement is SdtRow)
            {
                //skip
            }
            else if (lobjSdtElement is SdtRunRuby)
            {
                //skip
            }
            else if (lobjSdtElement is SdtRun)
            {
                SdtRun lobjSdtRun = (SdtRun)lobjSdtElement;
                SdtContentRun lobjSdtContentRun = lobjSdtElement.Descendants<SdtContentRun>().FirstOrDefault<SdtContentRun>();                
                if (lstrText.Contains(System.Environment.NewLine))
                {
                    // remove default text
                    lobjSdtContentRun.RemoveAllChildren<Run>();
                    // add text 
                    List<string> lcolValue = lstrText.Split(new string[] { System.Environment.NewLine }, StringSplitOptions.None).ToList<string>();
                    int i = 0;
                    foreach (string lstrValue in lcolValue)
                    {                        
                        Run lobjRun1 = new Run();
                        Text lobjText1 = new Text();
                        lobjText1.Text = lstrValue;
                        lobjRun1.Append(lobjText1);
                        lobjSdtContentRun.Append(lobjRun1);
                        if (i != lcolValue.Count - 1)
                        {
                            Run lobjRun2 = new Run();
                            Break break2 = new Break();
                            lobjRun2.Append(break2);
                            lobjSdtContentRun.Append(lobjRun2);
                        }
                        i++;
                    }                    
                }
                else
                {
                    Text t = lobjSdtElement.Descendants<Text>().FirstOrDefault();
                    if (t == null)
                    {                    
                        Run r = new Run();
                        t = new Text();
                        r.Append(t);
                        lobjSdtContentRun.Append(r);
                    }
                    t.Text = lstrText;
                }
            }
            else if (lobjSdtElement is SdtCell)
            {
                Text t = lobjSdtElement.Descendants<Text>().FirstOrDefault();
                if (t == null)
                {
                    Paragraph p = lobjSdtElement.Descendants<Paragraph>().FirstOrDefault();
                    if (p == null)
                    {
                        p = new Paragraph();
                    }
                    Run r = new Run();
                    t = new Text();
                    t.Text = lstrText;
                    r.Append(t);
                    p.Append(r);
                }
                t.Text = lstrText;
            }        
            return true;
        }

        #endregion

    }

}
