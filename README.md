FROM: https://github.com/shanselman/PowerPointToOBSSceneSwitcher

A .NET core based scene switcher than connects to OBS and changes scenes based note meta data. Put "OBS:Your Scene Name" on its own line in your notes and ensure the OBS Web Sockets Server is running and this app will change your scene as you change your PowerPoint slides.

Note this won't build with "dotnet build," instead open a Visual Studio 2019 Developer Command Prompt and build with "msbuild"