using GenAI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using System;
using System.Diagnostics;
using System.Windows.Forms;
using Utilities;
using static System.Net.Mime.MediaTypeNames;
using Task = System.Threading.Tasks.Task;


namespace AI_Composer
{
    public partial class GeminiComposer
    {
        CommandBarButton GeminiButton;
        private void ActivateGemini(object sender, System.EventArgs e)
        {
            string buttonName = "InnoWrite AI";
            try
            {
                CommandBar commandBar = Application.CommandBars["Text"];

                GeminiButton = FindButtonByName(commandBar, buttonName);
                if (GeminiButton == null)
                {
                    GeminiButton = (CommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
                    GeminiButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    GeminiButton.FaceId = 31;
                    GeminiButton.Caption = buttonName;
                }

                GeminiButton.Click += new _CommandBarButtonEvents_ClickEventHandler(GeminiComposer_Activate);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private CommandBarButton FindButtonByName(CommandBar commandBar, string buttonName)
        {
            foreach (CommandBarControl control in commandBar.Controls)
            {
                if (control is CommandBarButton button && control.Caption == buttonName)
                {
                    return button;
                }
            }
            return null;
        }

        private void GeminiComposer_Activate(CommandBarButton Button, ref bool CancelDefault)
        {
            if (Application.Selection == null || Application.Selection.Text.Trim() == String.Empty)
            {
                MessageBox.Show("Please select some text for the query context.", "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string apiKey = Environment.GetEnvironmentVariable("GEMINI_API_KEY");

            if (string.IsNullOrEmpty(apiKey))
            {
                DialogResult result = MessageBox.Show("Please set the API key before using, and remember to restart Microsoft Word after setting the API key.", "API key not found", MessageBoxButtons.OK, MessageBoxIcon.Error);

                if (result == DialogResult.OK)
                {
                    Process.Start("https://github.com/phanxuanquang/AI-Composer");
                }
                return;
            }

            try
            {
                var prompt = $"You are AI Composer, an expert in content composition with over 20 years of experience. Consider the topic and the request in my input, and compose the content accordingly. The input is: '{Application.Selection.Text}'. \nYour content:";
                var result = string.Empty;
                Task.Run(async () =>
                {
                    result = await Generator.GenerateContent(apiKey, prompt, false, CreativityLevel.Medium, GenerativeModel.Gemini_15_Flash);
                }).Wait();

                Application.Selection.Text = result;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error why composing content. Please try again after 1 minute.\nError: {ex.Message}.{ex.InnerException?.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region VSTO generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ActivateGemini);
        }

        #endregion
    }
}
