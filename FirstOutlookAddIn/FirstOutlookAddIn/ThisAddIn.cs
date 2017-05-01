using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows;
using System.Windows.Forms;
using Microsoft.ProjectOxford.Text.Sentiment;
using System.Drawing;

namespace FirstOutlookAddIn
{
    public partial class ThisAddIn
    {
        Outlook.Explorer currentExplorer = null;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            currentExplorer = this.Application.ActiveExplorer();
            currentExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(CurrentExplorer_Event);
        }
        private void CurrentExplorer_Event()
        {
            if (this.Application.ActiveExplorer().Selection.Count > 0)
            {
                Object selObject = this.Application.ActiveExplorer().Selection[1];
                if (selObject is Outlook.MailItem)
                {
                    Outlook.MailItem mailItem = (selObject as Outlook.MailItem);
                    var prop = mailItem.UserProperties.Find("Analyze", true);
                    if (prop == null)
                    {
                        prop = mailItem.UserProperties.Add("Analyze", Outlook.OlUserPropertyType.olText);
                    }
                    if (string.IsNullOrEmpty(prop.Value))
                    {
                        prop.Value = GetSentiment(mailItem.Body);
                    }
                    mailItem.Save();
                }
            }
        }
        private static string GetSentiment(string text)
        {
            try
            {
                double score = 0;
                var apiKey = "Tu api key de azure cognitive service text"";
                var document = new SentimentDocument()
                {
                    Id = "OutlookSergio",
                    Text = text,
                    Language = "es"
                };
                var request = new SentimentRequest();
                request.Documents.Add(document);
                var client = new SentimentClient(apiKey);
                var response = client.GetSentiment(request);
                foreach (var doc in response.Documents)
                {
                    score += doc.Score;
                }
                score = Math.Round((score / response.Documents.Count), 2) * 100;
                return score + "%";
            }
            catch(Exception e)
            {
                return null;
            }
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Nota: Outlook ya no genera este evento. Si tiene código que 
            //    deba ejecutarse cuando Outlook se cierre, vea http://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region Código generado por VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
