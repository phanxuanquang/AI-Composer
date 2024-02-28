using Markdig;
using Microsoft.Office.Core;
using Newtonsoft.Json;
using System;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AI_Composer
{
    public partial class GeminiComposer
    {
        private void ActivateGemini(object sender, System.EventArgs e)
        {
            Integrate_GeminiComposer();
        }

        private void Integrate_GeminiComposer()
        {
            try
            {
                MsoControlType buttonControlType = MsoControlType.msoControlButton;

                CommandBar commandBar = Application.CommandBars["Text"];
                CommandBarButton GeminiButton = (CommandBarButton)commandBar.Controls.Add(buttonControlType, Type.Missing, Type.Missing, Type.Missing, true);

                GeminiButton.Style = MsoButtonStyle.msoButtonIconAndCaption;
                GeminiButton.FaceId = 31;
                GeminiButton.Caption = "AI Composer";

                GeminiButton.Click += new _CommandBarButtonEvents_ClickEventHandler(GeminiComposer_Activate);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GeminiComposer_Activate(CommandBarButton Ctrl, ref bool CancelDefault)
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
                    Process.Start("https://github.com/phanxuanquang/Gemini-Writter");
                }
                return;
            }

            try
            {
                Input input = JsonConvert.DeserializeObject<Input>("{\"contents\":[{\"parts\":[{\"text\":\"Hello.\"}]}],\"safetySettings\":[{\"category\":\"HARM_CATEGORY_DANGEROUS_CONTENT\",\"threshold\":\"BLOCK_ONLY_HIGH\"}],\"generationConfig\":{\"stopSequences\":[\"Title\"],\"temperature\":0.5,\"maxOutputTokens\":4096,\"topP\":0.8,\"topK\":20}}");
                input.SetQuery($"Assume that you are a expert in copywriting field with over 20 years of experience. Considering the topic and the request in my input, help me to compose the content accordingly. The input is: '{Application.Selection.Text}'");

                Output output = null;
                Task.Run(async () =>
                {
                    output = await ComposeContentFrom(input, apiKey);
                }).Wait();

                if (output != null)
                {
                    Application.Selection.Text += "\n" + AsPlainText(output.candidates.FirstOrDefault().content.parts.FirstOrDefault().text);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error", ex.Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
