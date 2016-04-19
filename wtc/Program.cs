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

// wtc -o \\172.16.1.22\data\Vorlagen -n \\stor1.neos-it.local\data\Vorlagen -d c:\temp\wordtest -r


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


        [ParserState]
        public IParserState LastParserState { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            // 
            return "Word Template Corrector\nCorrecting wrong paths to templates in MS Office Word documents.\nUSE AT YOUR OWN RISK.\n\n" + 
                HelpText.AutoBuild(this, (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));
        }
    }

    class Program
    {
        static int Main(string[] args)
        {
            var options = new Options();
            if (CommandLine.Parser.Default.ParseArguments(args, options))
            {

                // check if folder exits
                if (!Directory.Exists(options.Directory))
                {
                    Console.WriteLine("Working directory does not exist.");
                    return 2;
                }

                // Initialize some variables
                string tempUnzipDirPrefix = "_wtc_";
                string tempDir = Path.GetTempPath();
                int fileCounter = 0; // counter for files
                int changeCounter = 0; // counter for corrected files
                int errorCounter = 0; // counter for errors
                bool error = false;
                bool changed = false;


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
                    changed = false;

                    Console.Write("         " + file);

                    string tempUnzipDir = tempDir + tempUnzipDirPrefix + Path.GetFileName(file);

                    // unzip
                    try
                    {

                        // unzip document to temp folder
                        ZipFile.ExtractToDirectory(file, tempUnzipDir);

                        string settingsFilePath = tempUnzipDir + @"\word\_rels\settings.xml.rels";
                        if (File.Exists(settingsFilePath))
                        {
                            string oldContent = File.ReadAllText(settingsFilePath);
                            string newContent = oldContent.Replace(options.Old, options.New); // replace
                            if (oldContent != newContent)
                            {
                                // check for DryRun
                                if (options.DryRun)
                                {
                                    changed = true;
                                    changeCounter++;
                                }
                                else
                                {

                                    File.WriteAllText(settingsFilePath, newContent);
                                    changed = true;
                                    changeCounter++;

                                    // save original file
                                    try
                                    {
                                        File.Move(file, file + ".bak");

                                        // Re-Zip files to docx
                                        try
                                        {
                                            ZipFile.CreateFromDirectory(tempUnzipDir, file);

                                            // delete backup file if wanted
                                            if (options.NoBackup)
                                            {
                                                File.Delete(file + ".bak");
                                            }
                                        }
                                        catch (Exception e2)
                                        {
                                            error = true;
                                            Console.Write(" - rezip failed: {0}", e2.Message);

                                            // undo rename
                                            File.Move(file + ".bak", file);
                                            Console.Write(" - backup restored");
                                        }
                                    }
                                    catch (Exception e3)
                                    {
                                        error = true;
                                        Console.Write(" - creating backup file failed: {0}", e3.Message);
                                    }
                                    finally { }
                                }
                            }
                        }
                        // remove unzipped files and temp folder
                        Directory.Delete(tempUnzipDir, true);
                    }
                    catch (Exception e1)
                    {
                        error = true;
                        Console.Write(" - an error occured: {0}", e1.Message);
                    }
                    finally { }

                    Console.Write("\r");
                    if (error == true)
                    {
                        errorCounter++;
                        Console.Write("FAILED");
                    }
                    else
                    {
                        if (changed == true)
                        {
                            if (options.DryRun)
                            {
                                Console.Write("AFFECTED");
                            } else {
                                Console.Write("CHANGED");
                            }
                        }
                    }
                    Console.Write("\n");

                }

                // Get the elapsed time as a TimeSpan value.
                stopWatch.Stop();
                TimeSpan ts = stopWatch.Elapsed;

                Console.WriteLine(fileCounter + " file(s) scanned");
                Console.Write(changeCounter + " file(s) ");
                if (options.DryRun)
                {
                    Console.WriteLine(" need correction");
                }else
                {
                    Console.WriteLine(" need changed");
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
    }
}
