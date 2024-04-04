using System;
using Microsoft.CognitiveServices.Speech;
using OpenQA.Selenium.Edge;

public class Profile
{
    public string Name { get; internal set; }
    public string TranscriptionName { get; internal set; }

    public string EdgeLink { get; internal set; }
    public EdgeDriver Driver { get; internal set; }
    public string SynthesizedVoice { get; internal set; }

    public SpeechConfig ProfileSpeechConfig { get; internal set; }
}