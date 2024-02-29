using Markdig;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using Task = System.Threading.Tasks.Task;


namespace AI_Composer
{
    public partial class GeminiComposer
    {
        CommandBarButton GeminiButton;
        private void ActivateGemini(object sender, System.EventArgs e)
        {
            string buttonName = "AI Composer";
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
                if (control is CommandBarButton && control.Caption == buttonName)
                {
                    return (CommandBarButton)control;
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
                Input input = JsonConvert.DeserializeObject<Input>("{\"contents\":[{\"parts\":[{\"text\":\"Hello.\"}]}],\"safetySettings\":[{\"category\":\"HARM_CATEGORY_DANGEROUS_CONTENT\",\"threshold\":\"BLOCK_ONLY_HIGH\"}],\"generationConfig\":{\"stopSequences\":[\"Title\"],\"temperature\":0.5,\"maxOutputTokens\":4096,\"topP\":0.8,\"topK\":20}}");
                input.SetQuery($"You are AI Composer, an expert in content composition with over 20 years of experience. Consider the topic and the request in my input, and compose the content accordingly. The input is: '{Application.Selection.Text}'");

                Output output = null;
                Task.Run(async () =>
                {
                    output = await ComposeContentFrom(input, apiKey);
                }).Wait();
                try
                {
                    Application.Selection.Text = AsPlainText(output.candidates.FirstOrDefault().content.parts.FirstOrDefault().text);
                }
                catch(Exception ex)
                {
                    MessageBox.Show($"Error why composing content. Please try again after 1 minute.\nError: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public async Task<Output> ComposeContentFrom(Input input, string apiKey)
        {
            string model = "gemini-1.0-pro";
            string apiUrl = $"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={apiKey}";

            using (var httpClient = new HttpClient())
            {
                using (var request = new HttpRequestMessage(new HttpMethod("POST"), apiUrl))
                {
                    request.Content = new StringContent(JsonConvert.SerializeObject(input));
                    request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                    try
                    {
                        var response = await httpClient.SendAsync(request);
                        var res = await response.Content.ReadAsStringAsync();
                        var output = JsonConvert.DeserializeObject<Output>(res);
                        return output;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return null;
                    }
                }
            }
        }

        public string AsPlainText(string markdown)
        {
            var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
            var result = Markdown.ToPlainText(markdown, pipeline);
            return result.Trim();
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
