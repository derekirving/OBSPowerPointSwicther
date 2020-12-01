Orginal idea from: https://github.com/shanselman/PowerPointToOBSSceneSwitcher

See this in action: [Video](https://web.microsoftstream.com/video/c3f4e61f-b595-4c74-8dcf-e37e92231fbf)

A .NET core based scene switcher than connects to OBS and changes scenes based note meta data. Put *"OBS:[Your-Scene-Name]"* on its own line in your notes and ensure the OBS Web Sockets Server is running and this app will change your scene as you change your PowerPoint slides.

Note this won't build with "dotnet build," open with Visual Studio 2019 and build from there.

There is a 7Zip archive in the [Assets/Powerpoint](./Assets/Powerpoint) folder that includes 2 presentations - one for the 4:3 Microsoft LifeCam and the other for a 16:9 DSLR
The example OBS Scene Collections are in [Assets/Obs](./Assets/Obs)
