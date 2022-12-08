using System;
using System.IO;
using System.Collections.Generic;
using CommandLine;

namespace Validate.Spreadsheet
{
    class Program
    {
        public class Options
        {
            [Option('i', "inputfilepath", Required = true, HelpText = "The input filepath.")]
            public string InputFilepath { get; set; }

            [Option('s', "standard", Required = false, HelpText = "Set to validate standard.")]
            public bool Standard { get; set; }

            [Option('a', "archivalrequirements", Required = false, HelpText = "Set to validate archival requirements.")]
            public bool ArchivalRequirements { get; set; }
        }

        public static int Main(string[] args)
        {
            // Parse user arguments
            var parser = new Parser(with => with.HelpWriter = null);
            int exitcode = parser.ParseArguments<Options>(args).MapResult((opts) => RunApp(opts), errs => ShowHelp(errs));
            return exitcode;
        }

        static int RunApp(Options arg)
        {
            int ExitCode, fail = 2;
            bool fileFormat_success = false, archivalReq_success = false;

            // Inform user of input filepath
            Console.WriteLine($"Validating: {arg.InputFilepath}");

            if (File.Exists(arg.InputFilepath))
            {
                try
                {
                    string extension = Path.GetExtension(arg.InputFilepath);
                    switch (extension.ToLower()) // The switch includes all accepted file extensions
                    {
                        case ".fods":
                        case ".ods":
                        case ".ots":
                            Validate_ODS ods = new Validate_ODS();

                            if (arg.Standard == true)
                            {
                                // Validate file format standard
                                fileFormat_success = ods.Validate_Standard(arg.InputFilepath);
                            }

                            if (arg.ArchivalRequirements == true)
                            {
                                // Validate archival requirements
                                archivalReq_success = ods.Validate_ArchivalRequirements(arg.InputFilepath);
                            }

                            // Return exit code
                            ExitCode = Program.ExitCode(arg, fileFormat_success, archivalReq_success);
                            return ExitCode;

                        case ".xlsb":
                            // Inform user validation is not possible
                            Console.WriteLine("Validation of .xlsb file format is not supported because of its binary structure");
                            return fail;

                        case ".xlsm":
                        case ".xlsx":
                        case ".xltm":
                        case ".xltx":
                            Validate_XLSX xlsx = new Validate_XLSX();

                            if (arg.Standard == true)
                            {
                                // Validate file format standard
                                fileFormat_success = xlsx.Validate_Standard(arg.InputFilepath);
                            }

                            if (arg.ArchivalRequirements == true)
                            {
                                // Validate archival requirements
                                archivalReq_success = xlsx.Validate_ArchivalRequirements(arg.InputFilepath);
                            }

                            // Return exit code
                            ExitCode = Program.ExitCode(arg, fileFormat_success, archivalReq_success);
                            return ExitCode;

                        default:
                            Console.WriteLine("File format is not an accepted file format");
                            return fail;
                    }
                }
                // If spreadsheet is password protected or otherwise unreadable
                catch (System.FormatException) 
                {
                    Console.WriteLine("File cannot be read");
                    return fail;
                }
                catch (System.IO.InvalidDataException)
                {
                    Console.WriteLine("File cannot be read");
                    return fail;
                }
                catch(DocumentFormat.OpenXml.Packaging.OpenXmlPackageException)
                {
                    Console.WriteLine("File cannot be read");
                    return fail;
                }
            }
            // If filepath has no file
            else
            {
                Console.WriteLine("No file in filepath");
                return fail;
            }
        }

        // Provide exit code
        static int ExitCode(Options arg, bool fileFormat_success, bool archivalReq_success)
        {
            int invalid = 0, valid = 1, fail = 2;
            if (arg.Standard == true && arg.ArchivalRequirements == true)
            {
                if (fileFormat_success == false || archivalReq_success == false)
                {
                    return invalid;
                }
                else if (fileFormat_success == true && archivalReq_success == true)
                {
                    return valid;
                }
            }
            else if (arg.Standard == true && arg.ArchivalRequirements == false)
            {
                if (fileFormat_success == false)
                {
                    return invalid;
                }
                else if (fileFormat_success == true)
                {
                    return valid;
                }
            }
            else if (arg.Standard == false && arg.ArchivalRequirements == true)
            {
                if (archivalReq_success == false)
                {
                    return invalid;
                }
                else if (archivalReq_success == true)
                {
                    return valid;
                }
            }
            Console.WriteLine("No validation arguments were input");
            return fail;
        }

        // Show help to user, if parsing arguments fail
        static int ShowHelp(IEnumerable<Error> errs)
        {
            int fail = 2;
            Console.WriteLine("Input arguments have errors");
            return fail;
        }
    }
}