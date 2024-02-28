using System.Collections.Generic;

namespace AI_Composer
{
    public class Output
    {
        public List<Candidate> candidates { get; set; }
        public PromptFeedback promptFeedback { get; set; }
    }
    public class Candidate
    {
        public OutputContent content { get; set; }
        public string finishReason { get; set; }
        public int? index { get; set; }
        public List<SafetyRating> safetyRatings { get; set; }
    }

    public class OutputContent
    {
        public List<OutputPart> parts { get; set; }
        public string role { get; set; }
    }

    public class OutputPart
    {
        public string text { get; set; }


    }

    public class PromptFeedback
    {
        public List<SafetyRating> safetyRatings { get; set; }
    }
    public class SafetyRating
    {
        public string category { get; set; }
        public string probability { get; set; }
    }
}
