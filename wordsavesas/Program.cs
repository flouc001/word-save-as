using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Word;
using CommandLine;
using CommandLine.Text;

namespace wordsaveas
{
    class Program
    {
        class Options {
            [Option('i', "input", Required = true, HelpText = "Input file/directory to be processed.")]
            public string SourcePath { get; set; }

            [Option('o', "output", Required = true, HelpText = "Output file/directory to receive processed files")]
            public string DestPath { get; set; }

            [Option("format", Required = false, HelpText = "Output format :: pdf, text, doc.", DefaultValue = "pdf")]
            public string OutputFormat { get; set; }

            [Option('f',"force", Required = false, HelpText = "Force file overwrite if output path already exists.", DefaultValue = false)]
            public bool ForceDelete { get; set; }

            [Option('m', "mask", Required = false, HelpText = "File mask to filter input files default: \"*.docx\"", DefaultValue = "*.docx")]
            public string FileMask { get; set; }

            [ParserState]
            public IParserState LastParserState { get; set; }

            [HelpOption]
            public string GetUsage() {
                return HelpText.AutoBuild(this, (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));
            }
        }

        static void SaveAsFormat(String src, String dest, Application wi, WdSaveFormat fmt, bool del) {
            if (File.Exists(dest) && del) File.Delete(dest);
            else if (File.Exists(dest)) { throw new Exception("Output file already exists and force parameter is false."); }
            wi.Documents.Open(src);
            wi.ActiveDocument.SaveAs2(dest, fmt);
            Console.WriteLine($"{src} -> {dest}");
            wi.ActiveDocument.Close();
        }

        static bool IsDirectory(String path) {
            FileAttributes iattr = File.GetAttributes(path);
            if ((iattr & FileAttributes.Directory) == FileAttributes.Directory) return true;
            return false;
        }

        static WdSaveFormat SaveFormat(string fmt) {
            switch (fmt.ToLower())
            {
                case "pdf": return WdSaveFormat.wdFormatPDF;
                case "doc": return WdSaveFormat.wdFormatDocument97;
                case "html": return WdSaveFormat.wdFormatFilteredHTML;
                case "txt": case "text": return WdSaveFormat.wdFormatText;
            }
            throw new Exception("Format not recognized");
        }

        static void Main(string[] args)
        {
            Options opts = new Options();
            ConsoleColor col = Console.ForegroundColor; // Get default color
            if (CommandLine.Parser.Default.ParseArguments(args, opts)) {
                string input = opts.SourcePath, output = opts.DestPath;
                if (!File.Exists(input) && !Directory.Exists(input)) throw new Exception("Input file/directory does not exist.");
                bool dir = IsDirectory(input);
                if (dir && !Directory.Exists(output)) { throw new Exception("Output is not a directory"); }
                Console.Write("Initialising..");
                Application wi = new Application(); // Create new Word instance
                try
                {
                    Console.Write(".\n");
                    if (dir)
                    {
                        DirectoryInfo input_directory = new DirectoryInfo(input);
                        FileInfo[] files = input_directory.GetFiles(opts.FileMask);
                        foreach (FileInfo f in files)
                        {
                            string dest_name = $"{output}\\{f.Name.Substring(0, f.Name.LastIndexOf('.'))}.{opts.OutputFormat}";
                            if (!IsDirectory(f.FullName)) SaveAsFormat(f.FullName, dest_name, wi, SaveFormat(opts.OutputFormat), opts.ForceDelete);
                        }
                    }
                    else SaveAsFormat(input, output, wi, SaveFormat(opts.OutputFormat), opts.ForceDelete);
                }
                catch (Exception ex) {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(ex.Message);
                }
                wi.Quit();
            }
            Console.ForegroundColor = col;
            Console.WriteLine("Done!");
        }
    }
}
