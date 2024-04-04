# Summary
This app is intended to run various instances of Microsoft Teams in different browser profiles to automate the recording of a meeting transcription that can be used to demo Copilot. This will allow different types of meeting to be created without having to bring together groups of people.

# Minimal path to awesome

The following things are needed to make this script work:
- A set of accounts in Microsoft 365 with at least one that has a Copilot license
- Each account set up with an Edge profile and logged in to that profile - see https://support.microsoft.com/en-us/topic/sign-in-and-create-multiple-profiles-in-microsoft-edge-df94e622-2061-49ae-ad1d-6f0e43ce6435 for more information on Edge profiles.
- Set up a meeting with one of the accounts and copy the link
- Set up an Azure Speech Service and get the key and region


The following changes are needed

- Extract the meeting ID from the meeting URL and replace in the meetingPrefix variable
- Extract the querystrings from the meeting link and add them to the meetingSuffix
- Set your speech key and region in environment variables with the commands below (or your preferred way of setting environment variables)
    - setx SPEECH_KEY "XXXXXXX"
    - setx SPEECH_REGION "westus"

The following change is optional
- Use your own prompt to create a transcription and copy the text into the Transcript.txt file.
- Change the details within the profiles collection to set the specific profile, name and even speech synth to use.

Prior to running the script, if you use an external microphone, position it near to the speaker. You may need to try out different configurations of microphone and speaker, testing how well it is working with Live Captions enabled to see what is being said.

Once complete, execute dotnet run.

# Architectural overview

This uses Selenium to automate the Edge browsers and Azure Speech Service to read out the transcription.

# Sample prompt to create a meeting transcript:

Generate a transcript for a Teams meeting between two people with the following criteria in Copilot in Word:
- There are two people attending this project status update meeting
- Kevin is the first person attending and is an established project manager who is feeling stressed about meeting deadlines. He has been working in the industry many years but this is the first project that he has been truly challenged by
- Bluey is the second person attending the meeting and is a blue cartoon dog that is far too young to be on this project but for some reason she is. She has a bad attitude to work and is not helpful
- This meeting should cover what the plans were for the project, what the current status is and what the next steps are
- There should be at least 8 interruptions between the two attendees
- There should be no clear outcome