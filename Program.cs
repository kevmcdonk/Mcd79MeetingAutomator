using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.CognitiveServices.Speech;
using Microsoft.CognitiveServices.Speech.Audio;
using System.Text.RegularExpressions;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;

class Program
{
    // This example requires environment variables named "SPEECH_KEY" and "SPEECH_REGION"
    static string speechKey = Environment.GetEnvironmentVariable("SPEECH_KEY");
    static string speechRegion = Environment.GetEnvironmentVariable("SPEECH_REGION");

    string meetingPrefix = "MDE1NmGmNWItO2FkMi12NjE0LVkyNDYtYTI5OWR0ZWY3MWUx";
    string meetingSuffix="0?context=%7b%22Tid%22%3a%22b618d3d2-c131-42fd-a629-3711d84af275%22%2c%22Oid%22%3a%2258e4a1fb-5d3g-4g9d-b210-e9e9d1437a2b%22%7d&anon=true";
    static string meetingLink = $"https://teams.microsoft.com/_#/l/meetup-join/19:meeting_{meetingPrefix}@thread.v2/${meetingSuffix}";


    static void OutputSpeechSynthesisResult(SpeechSynthesisResult speechSynthesisResult, string text)
    {
        switch (speechSynthesisResult.Reason)
        {
            case ResultReason.SynthesizingAudioCompleted:
                Console.WriteLine($"Speech synthesized for text: [{text}]");
                break;
            case ResultReason.Canceled:
                var cancellation = SpeechSynthesisCancellationDetails.FromResult(speechSynthesisResult);
                Console.WriteLine($"CANCELED: Reason={cancellation.Reason}");

                if (cancellation.Reason == CancellationReason.Error)
                {
                    Console.WriteLine($"CANCELED: ErrorCode={cancellation.ErrorCode}");
                    Console.WriteLine($"CANCELED: ErrorDetails=[{cancellation.ErrorDetails}]");
                    Console.WriteLine($"CANCELED: Did you set the speech resource key and region values?");
                }
                break;
            default:
                break;
        }
    }

    static List<Profile> SetupBrowsers()
    {
        List<IWebDriver> drivers = new List<IWebDriver>();
        List<Profile> profiles = new List<Profile>();
        profiles.Add(new Profile { 
            EdgeLink = "C:\\Users\\firstname.lastname\\AppData\\Local\\Microsoft\\Edge\\User Data\\Profile 1",
            Name = "Kevin", 
            TranscriptionName = "Kevin", 
            SynthesizedVoice = "en-GB-ThomasNeural"
        });
        profiles.Add(new Profile { 
            EdgeLink = "C:\\Users\\firstname.lastname\\AppData\\Local\\Microsoft\\Edge\\User Data\\Profile 2",
            Name = "Bluey",
            TranscriptionName = "Bluey",
            SynthesizedVoice = "en-AU-NatashaNeural"
        });

        foreach (var profile in profiles)
        {
            var options = new EdgeOptions();
            options.AddArgument($"user-data-dir={profile.EdgeLink}");
            options.AddArgument("--mute-audio");
            EdgeDriver driver = new EdgeDriver(options);
            profile.Driver = driver;
        }
        return profiles;
    }

    static void TeardownBrowsers(List<Profile> profiles)
    {
        foreach (var profile in profiles)
        {
            profile.Driver.Close();
        }
    }

    public static async void StartMeetings(List<Profile> profiles)
    {
        foreach (var profile in profiles)
        {
            var driver = profile.Driver;
            driver.Navigate().GoToUrl(meetingLink);
            // Need to sign in first! Not handled in code

            // Work out how to handle speaker options...
            // Work out how to handle "Allow microphone"
            IWebElement iframe = null;

            bool elementLoading = true;
            bool iFrameFound = false;
            int waitCountLimit = 4;
            int waitCount = 0;
            while (elementLoading && waitCount < waitCountLimit)
            {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
                try
                {
                    waitCount++;
                    iframe = driver.FindElement(By.ClassName("embedded-page-content"));
                    elementLoading = false;
                    iFrameFound = true;
                }
                catch (Exception ex)
                {
                    elementLoading = true;
                }
            }

            if (iFrameFound)
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(2));
                wait.Until(d => iframe.Displayed);
                //Switch to the frame
                driver.SwitchTo().Frame(iframe);
            }

            //Now we can click the button
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(15);
            IWebElement button = driver.FindElement(By.XPath("//*[@id=\"prejoin-join-button\"]"));
            WebDriverWait buttonWait = new WebDriverWait(driver, TimeSpan.FromSeconds(2));
            buttonWait.Until(d => button.Displayed);
            button.Click();

            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(15);
            WebDriverWait muteButtonWait = new WebDriverWait(driver, TimeSpan.FromSeconds(2));
            IWebElement muteButton = driver.FindElement(By.Id("microphone-button"));
            muteButtonWait.Until(d => muteButton.Displayed);
            string muteButtonState = muteButton.GetAttribute("data-state");
            if (muteButtonState == "mic") // can be mic or mic-off
            {
                muteButton.Click();
            }
        }
    }

    async static Task ReadTranscript(List<Profile> profiles)
    {
        foreach (var profile in profiles)
        {
            var speechConfig = SpeechConfig.FromSubscription(speechKey, speechRegion);
            speechConfig.SpeechSynthesisVoiceName = profile.SynthesizedVoice;
            profile.ProfileSpeechConfig = speechConfig;

        }


        using (var sr = new StreamReader("Transcript.txt"))
        {
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                Console.WriteLine(line);
                string pattern = @"\[(\d{2}:\d{2})\] (\w+): (.*)";
                Match match = Regex.Match(line, pattern);

                if (match.Success)
                {
                    string time = match.Groups[1].Value;
                    string name = match.Groups[2].Value;
                    string text = match.Groups[3].Value;

                    // Console.WriteLine($"Time: {time}");
                    // Console.WriteLine($"Name: {name}");
                    // Console.WriteLine($"Text: {text}");

                    foreach (var profile in profiles)
                    {
                        Console.WriteLine($"Profilename: {profile.TranscriptionName}");
                        if (name == profile.TranscriptionName)
                        {
                            Console.WriteLine("Speaking");
                            profile.Driver.SwitchTo();
                            WebDriverWait muteButtonWait = new WebDriverWait(profile.Driver, TimeSpan.FromSeconds(2));
                            IWebElement muteButton = profile.Driver.FindElement(By.Id("microphone-button"));
                            muteButtonWait.Until(d => muteButton.Displayed);
                            string muteButtonState = muteButton.GetAttribute("data-state");
                            if (muteButtonState == "mic-off") // can be mic or mic-off
                            {
                                muteButton.Click();
                            }
                            using (var synthesizer = new SpeechSynthesizer(profile.ProfileSpeechConfig)) {
                                var speechSynthesisResult = await synthesizer.SpeakTextAsync(text);
                                OutputSpeechSynthesisResult(speechSynthesisResult, text);
                                Console.WriteLine("Speech finished");
                            }
                            muteButton = profile.Driver.FindElement(By.Id("microphone-button"));
                            muteButtonWait.Until(d => muteButton.Displayed);
                            muteButtonState = muteButton.GetAttribute("data-state");
                            if (muteButtonState == "mic") // can be mic or mic-off
                            {
                                muteButton.Click();
                            }
                        }

                    }
                }
            }
        }
    }


    async static Task Main(string[] args)
    {
        var profiles = SetupBrowsers();
        StartMeetings(profiles);
        await ReadTranscript(profiles);
        TeardownBrowsers(profiles);

        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }
}