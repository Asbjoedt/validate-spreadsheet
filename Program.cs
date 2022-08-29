using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

class Program
{
    public static string new_filepath = "";

    bool? Main(string[] args)
    {
        string filepath = args[0];
        string extension = Path.GetExtension(filepath);
        bool? success;
        try
        {
            switch (extension) // The switch includes all accepted file extensions
            {
                case ".fods":
                case ".ods":
                case ".ots":
                case ".FODS":
                case ".ODS":
                case ".OTS":
                    Console.WriteLine(filepath); // Write filepath to user
                    success = Validate_XLSX(filepath); // Convert data
                    return success;

                case ".xlsb":
                case ".xlsm":
                case ".xlsx":
                case ".xltm":
                case ".xltx":
                case ".XLSB":
                case ".XLSM":
                case ".XLSX":
                case ".XLTM":
                case ".XLTX":
                    Console.WriteLine(filepath); // Write filepath to user
                    success = Validate_XLSX(filepath); // Convert data
                    return success;

                default:
                    success = null;
                    Console.WriteLine("File format is not an accepted file format");
                    return success;
            }
        }
        catch (FileNotFoundException) // If filepath has not file
        {
            Console.WriteLine("No file in filepath");
            success = null;
            return success;
        }
        catch (FormatException) // If spreadsheet is password protected or otherwise unreadable
        {
            Console.WriteLine("File cannot be read");
            success = null;
            return success;
        }
    }

    static bool Validate_XLSX(string filepath)
    {
        bool success = false;
        using (var spreadsheet = SpreadsheetDocument.Open(filepath, false))
        {
            // Validate
            var validator = new OpenXmlValidator();
            var validation_errors = validator.Validate(spreadsheet).ToList();
            int error_count = validation_errors.Count;
            int error_number = 0;
            if (validation_errors.Any()) // If errors, inform user & return results
            {
                Console.WriteLine($"--> File format is invalid - Spreadsheet has {error_count} validation errors");
                foreach (var error in validation_errors)
                {
                    error_number++;
                    Console.WriteLine("--> Error " + error_number);
                    Console.WriteLine("----> Id: " + error.Id);
                    Console.WriteLine("----> Description: " + error.Description);
                    Console.WriteLine("----> Error type: " + error.ErrorType);
                    Console.WriteLine("----> Node: " + error.Node);
                    Console.WriteLine("----> Path: " + error.Path.XPath);
                    Console.WriteLine("----> Part: " + error.Part.Uri);
                    if (error.RelatedNode != null)
                    {
                        Console.WriteLine("----> Related Node: " + error.RelatedNode);
                        Console.WriteLine("----> Related Node Inner Text: " + error.RelatedNode.InnerText);
                    }
                }
                success = false;
            }
            else
            {
                Console.WriteLine($"--> File format is valid");
                success = true;
            }
        }
        return success;
    }

    static bool Validate_ODS(string filepath)
    {
        bool success = false;
        Process app = new Process();
        string? dir = null;
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) // If app is run on Windows
        {
            dir = Environment.GetEnvironmentVariable("ODFValidator");
        }
        if (dir != null)
        {
            app.StartInfo.FileName = dir;
        }
        else
        {
            app.StartInfo.FileName = "C:\\Program Files\\Beyond Compare 4\\BCompare.exe";
        }
        app.StartInfo.Arguments = $"java -jar \"{filepath}/odfvalidator-VERSION-jar-with-dependencies.jar\"";
        app.Start();
        app.WaitForExit();
        int return_code = app.ExitCode;
        if (return_code == 0)
        {
            success = true;
        }
        app.Close();
        return success;
    }
}