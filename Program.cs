﻿using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;

namespace obscontrol
{
    class Program
    {
        private static Application ppt = new Microsoft.Office.Interop.PowerPoint.Application();
        private static ObsLocal OBS;
        static async Task Main(string[] args)
        {
            Console.Write("Connecting to OBS...");
            OBS = new ObsLocal();
            await OBS.Connect();
            Console.WriteLine("connected");
            // OBS.ChangeScene("Scene");
            // OBS.StartRecording();
            // Console.WriteLine("Recording...");
            // await Task.Delay(5000);
            // OBS.ChangeScene("Desktop");
            // await Task.Delay(5000);
            // OBS.StopRecording();
            // Console.WriteLine("Done...");

             Console.Write("Connecting to PowerPoint...");
            ppt.SlideShowNextSlide += App_SlideShowNextSlide;
            Console.WriteLine("connected to powerpoint");

            Console.ReadLine();
        }

        async static void App_SlideShowNextSlide(SlideShowWindow Wn)
        {
            if (Wn != null)
            {
                Console.WriteLine($"Moved to Slide Number {Wn.View.Slide.SlideNumber}");
                //Text starts at Index 2 ¯\_(ツ)_/¯
                var note = String.Empty;
                try { note = Wn.View.Slide.NotesPage.Shapes[2].TextFrame.TextRange.Text; }
                catch { /*no notes*/ }

                var notereader = new StringReader(note);
                string line;
                while ((line = notereader.ReadLine()) != null)
                {
                    if (line.StartsWith("OBS:")) {
                        line = line.Substring(4).Trim();
                        await HandleCommand(line);
                    }
                }
            }
        }

        static async Task HandleCommand(string command)
        {
            switch (command)
            {
                case "":
                    break;
                case "**START":
                    OBS.StartRecording();
                    break;
                case "**STOP":
                    OBS.StopRecording();
                    break;

                default:
                    Console.WriteLine($"  Switching to OBS Scene named \"{command}\"");
                    try { OBS.ChangeScene(command); }
                    catch (Exception ex) { Console.WriteLine($"  ERROR: {ex.Message.ToString()}"); }
                    break;
            }
        }
    }
}
