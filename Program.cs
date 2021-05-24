using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;
using Microsoft.WindowsAPICodePack.Shell;

namespace MetaDataWavExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            //ensure only wav files are in the directory
            string samplesDirectory;
            Console.WriteLine("Enter the file path:");
            samplesDirectory = (Console.ReadLine());
            string formattedDirectory = samplesDirectory.Replace("\""," ");
            CreateCSV(formattedDirectory);
        }

        static void CreateCSV(string formattedDirectory)
        {
            ulong duration;
            using (var stream = new StreamWriter("samplesinfo.csv", false, Encoding.UTF8))
            {
                //add table headers here
                var firstLine = string.Format("{0},{1},{2},{3}", "FileName", "Duration", "ChannelCount", "SampleRate");
                stream.WriteLine(firstLine);

                foreach (var f in new DirectoryInfo(formattedDirectory).GetFiles())
                {
                    using (var so = ShellObject.FromParsingName(f.FullName))
                    {
                        //add additional metadata to extract here
                        var fileName = so.Properties.GetProperty(SystemProperties.System.FileName).ValueAsObject.ToString();
                        IShellProperty dur = so.Properties.System.Media.Duration;
                        var t = (ulong)dur.ValueAsObject;
                        duration = (t / 10000000);
                        var channelCount = so.Properties.GetProperty(SystemProperties.System.Audio.ChannelCount).ValueAsObject.ToString();
                        var sampleRate = so.Properties.GetProperty(SystemProperties.System.Audio.SampleRate).ValueAsObject.ToString();

                        var line = string.Format("{0},{1},{2},{3}", fileName, duration, channelCount, sampleRate);
                        stream.WriteLine(line);
                    }
                }
            }
        }
    }
}
