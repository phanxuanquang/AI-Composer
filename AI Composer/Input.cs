using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AI_Composer
{
    public class Input
    {
        public List<InputContent> contents { get; set; }
        public List<SafetySetting> safetySettings { get; set; }
        public GenerationConfig generationConfig { get; set; }
        public void SetQuery(string query)
        {
            this.contents.FirstOrDefault().parts.FirstOrDefault().text = query;
        }
    }
    public class InputContent
    {
        public List<InputPart> parts { get; set; }
    }

    public class GenerationConfig
    {
        public List<string> stopSequences { get; set; }
        public double? temperature { get; set; }
        public int? maxOutputTokens { get; set; }
        public double? topP { get; set; }
        public int? topK { get; set; }
    }

    public class InputPart
    {
        public string text { get; set; }
    }

    public class SafetySetting
    {
        public string category { get; set; }
        public string threshold { get; set; }
    }
}
