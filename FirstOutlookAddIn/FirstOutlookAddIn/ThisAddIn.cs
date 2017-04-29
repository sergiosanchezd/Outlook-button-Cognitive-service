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

namespace FirstOutlookAddIn
{
    public partial class ThisAddIn
    {
        Office.CommandBar newToolBar;
        Office.CommandBarButton firstButton;
        Outlook.Explorers selectExplorers;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            selectExplorers = this.Application.Explorers;
            selectExplorers.NewExplorer += new Outlook.ExplorersEvents_NewExplorerEventHandler(newExplorer_Event);
            AddToolbar();
        }
        private void newExplorer_Event(Outlook.Explorer new_Explorer)
        {
            ((Outlook._Explorer)new_Explorer).Activate();
            newToolBar = null;
            AddToolbar();
        }

        private void AddToolbar()
        {
            if (newToolBar == null)
            {
                Office.CommandBars cmdBars = this.Application.ActiveExplorer().CommandBars;
                newToolBar = cmdBars.Add("Analyze text", Office.MsoBarPosition.msoBarTop, false, true);
            }
            try
            {
                Office.CommandBarButton button_1 = (Office.CommandBarButton)newToolBar.Controls.Add(1, missing, missing, missing, missing);
                button_1.Tag = "Button1";
                button_1.Caption = "Analyze text";
                button_1.Style = Office.MsoButtonStyle.msoButtonCaption;
                newToolBar.Visible = true;
                if (this.firstButton == null)
                {
                    this.firstButton = button_1;
                    firstButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(ButtonClick);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        } private void ButtonClick(Office.CommandBarButton ctrl, ref bool cancel)
        {
            if (this.Application.ActiveExplorer().Selection.Count > 0)
            {
                Object selObject = this.Application.ActiveExplorer().Selection[1];
                if (selObject is Outlook.MailItem)
                {
                    Outlook.MailItem mailItem = (selObject as Outlook.MailItem);
                    MainAsync(mailItem.Body);
                }
            }
        }
        static void MainAsync(string text)
        {
            double score = 0;
            var apiKey = "Tu api key del servicio de azure va aqui";
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
            MessageBox.Show("Score: " + score + "%");
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
