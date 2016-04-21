using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;
using System.Diagnostics;
using CommandLine;
using CommandLine.Text;

// wtc -d \\server\share\documents -o \\oldserver\share\templates\ -n \\server\share\templates\ -r


namespace WTC
{

    class Options
    {
        [Option('d', "directory", Required = true, HelpText = "working directory.")]
        public string Directory { get; set; }

        [Option('o', "old", Required = true, HelpText = "The old part of the templates path to be replaced.")]
        public string Old { get; set; }

        [Option('n', "new", Required = true, HelpText = "The new (replacement) part of the templates path.")]

        public string New { get; set; }

        [Option('r', "recursive", HelpText = "Recurse through subdirectories.")]
        public bool Recursive { get; set; }

        [Option('b', "nobackup", DefaultValue = false, HelpText = "Do NOT create a backup (.bak) of each changed document.")]
        public bool NoBackup { get; set; }

        [Option('t', "dry-run", DefaultValue = false, HelpText = "Do not change any files (for testing).")]
        public bool DryRun { get; set; }

        [Option('v', "verbose", DefaultValue = false, HelpText = "Activates verbose error messages.")]
        public bool Verbose { get; set; }


        [ParserState]
        public IParserState LastParserState { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            // 
            return "Word Template Corrector\nCorrecting wrong paths to templates in MS Office Word documents.\nUSE AT YOUR OWN RISK.\n\n" + 
                   HelpText.AutoBuild(this, (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current)) + "\n" +
                   @"  Example: wtc -d \\server\share\documents -o \\oldserver\share\templates\ -n \\server\share\templates\ -r" + "\n";
        }
    }


    class Program
    {

        public static string tempDir = Path.GetTempPath() + "_wtc_";

        static int Main(string[] args)
        {

            // Initialize some variables
            int fileCounter = 0; // counter for files
            int changeCounter = 0; // counter for corrected files
            int errorCounter = 0; // counter for errors
            int line; // for saving cursor Position
            ConsoleColor fgColor = Console.ForegroundColor;
            bool error = false;
            bool changed = false;

            var options = new Options();
            if (CommandLine.Parser.Default.ParseArguments(args, options))
            {

                // check if folder exits
                if (!Directory.Exists(options.Directory))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Working directory does not exist.");
                    Console.ForegroundColor = fgColor;
                    return 2;
                }

                // Output some information
                Console.WriteLine("Directory   : " + options.Directory);
                Console.WriteLine("Search for  : " + options.Old);
                Console.WriteLine("Replace with: " + options.New);
                Console.WriteLine("no Backups  : " + options.NoBackup.ToString());
                Console.WriteLine("Recursive   : " + options.Recursive.ToString());
                Console.WriteLine("Dry run     : " + options.DryRun.ToString());


                // start time measurement
                Stopwatch stopWatch = new Stopwatch();
                stopWatch.Start();

                var so = SearchOption.TopDirectoryOnly;
                if (options.Recursive)
                {
                    so = SearchOption.AllDirectories;
                }


                // fetch all possible affected documents
                var files = Directory.EnumerateFiles(options.Directory, "*.*", so)
                    .Where(s => s.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)
                             || s.EndsWith(".docm", StringComparison.OrdinalIgnoreCase)
                             || s.EndsWith(".docm", StringComparison.OrdinalIgnoreCase)
                             || s.EndsWith(".dotm", StringComparison.OrdinalIgnoreCase));


                // iterate through documents
                foreach (string file in files)
                {
                    fileCounter++;
                    error = false;

                    line = Console.CursorTop;
                    Console.Write("         " + file);


                    // lets try to correct the document
                    try
                    {
                        changed = correctDocument(file, options.Old, options.New, options.NoBackup, options.DryRun);
                    }
                    catch (Exception e)
                    {
                        error = true;
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.Write(" - error: {0}", e.Message);
                        if ((e.InnerException != null) & (options.Verbose))
                        {
                            Console.Write(" ({0})", e.InnerException.Message);
                            line = Console.CursorTop;
                        }
                        Console.ForegroundColor = fgColor;
                    }


                    Console.Write("\r");
                    if (!Console.IsOutputRedirected) { 
                        Console.CursorTop = line;
                    }

                    if (error == true)
                    {
                        errorCounter++;
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.Write("FAILED");
                        Console.ForegroundColor = fgColor;
                    }
                    else
                    {
                        if (changed == true)
                        {
                            changeCounter++;

                            if (options.DryRun)
                            {
                                Console.ForegroundColor = ConsoleColor.Yellow;
                                Console.Write("AFFECTED");
                            }
                            else {
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.Write("CHANGED");
                            }
                            Console.ForegroundColor = fgColor;
                        }
                    }
                    Console.Write("\n");

                }

                // Output results

                // Get the elapsed time as a TimeSpan value.
                stopWatch.Stop();
                TimeSpan ts = stopWatch.Elapsed;

                Console.WriteLine(fileCounter + " file(s) scanned");
                Console.Write(changeCounter + " file(s) ");
                if (options.DryRun)
                {
                    Console.WriteLine("affected and need correction");
                }else
                {
                    Console.WriteLine("corrected");
                }
                Console.WriteLine(errorCounter + " error(s) occured");


                // Format and display the TimeSpan value.
                string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                    ts.Hours, ts.Minutes, ts.Seconds,
                    ts.Milliseconds / 10);
                Console.WriteLine("Runtime " + elapsedTime);

                return 0;
            }
            else
            {
                return 1;
            }
        }

        /// <summary>
        /// Corrects the template path in a specific word document
        /// </summary>
        /// <param name="file">path to document in filesystem</param>
        /// <param name="oldPath">old path to template in document</param>
        /// <param name="newPath">new path to template in document</param>
        /// <param name="makeBackup">create backup file for every corrected document</param>
        /// <param name="dryRun">if true the original file will not be changed</param>
        /// <returns>file is changed or affected</returns>
        static bool correctDocument(string file, string oldPath, string newPath,  bool noBackup, bool dryRun)
        {

            bool changed = false;

            string tempUnzipDir = tempDir + Path.GetFileName(file);

            // unzip
            try
            {
                // unzip document to temp folder
                ZipFile.ExtractToDirectory(file, tempUnzipDir);

                string settingsFilePath = tempUnzipDir + @"\word\_rels\settings.xml.rels";
                if (File.Exists(settingsFilePath))
                {
                    string oldContent = File.ReadAllText(settingsFilePath);
                    string newContent = oldContent.Replace(oldPath, newPath); // replace

                    // something changed?
                    if (oldContent != newContent)
                    {
                        // check for DryRun
                        if (dryRun)
                        {
                            changed = true;
                        }
                        else
                        {

                            File.WriteAllText(settingsFilePath, newContent);
                            changed = true;

                            // save original file
                            try
                            {
                                File.Move(file, file + ".bak");

                                // Re-Zip files to docx
                                try
                                {
                                    ZipFile.CreateFromDirectory(tempUnzipDir, file);
                                    
                                    // delete backup file if wanted
                                    if (noBackup)
                                    {
                                        File.Delete(file + ".bak");
                                    }
                                }
                                catch (Exception e2)
                                {
                                    // undo rename
                                    try
                                    {
                                        File.Move(file + ".bak", file);
                                    }
                                    catch (Exception e4)
                                    {
                                        WTCException wtcEx4 = new WTCException("failed to remove backup file", e4);
                                        throw wtcEx4;
                                    }

                                    WTCException wtcEx2 = new WTCException("failed to rezip", e2);
                                    throw wtcEx2;
                                }
                            }
                            catch (Exception e3)
                            {
                                WTCException wtcEx3 = new WTCException("failed to create backup file", e3);
                                throw wtcEx3;
                            }
                        }
                    }
                }

                // remove unzipped files and temp folder
                try
                {
                    Directory.Delete(tempUnzipDir, true);
                }
                catch (Exception e5)
                {
                    WTCException wtcEx5 = new WTCException("failed to remove temporary unzip directory", e5);
                    throw wtcEx5;
                }

            }
            catch (Exception e1)
            {
                WTCException wtcEx1 = new WTCException("failed to unzip document", e1);
                throw wtcEx1;
            }
            return changed;
        }


        [Serializable()]
        public class WTCException : System.Exception
        {
            public WTCException() : base() { }
            public WTCException(string message) : base(message) { }
            public WTCException(string message, System.Exception inner) : base(message, inner) { }

            // A constructor is needed for serialization when an
            // exception propagates from a remoting server to the client. 
            protected WTCException(System.Runtime.Serialization.SerializationInfo info,
                System.Runtime.Serialization.StreamingContext context)
            { }
        }

    }
}
