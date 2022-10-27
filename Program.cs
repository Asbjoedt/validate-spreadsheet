using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace Validate.Spreadsheet
{
    class Program
    {
        public static int Main(string[] args)
        {
            string filepath = args[0];
            string extension = System.IO.Path.GetExtension(filepath);
            int invalid = 0;
            int valid = 1;
            int fail = 2;
            bool fileFormat_success = false;
            bool archivalReq_success = false;

            try
            {
                switch (extension) // The switch includes all accepted file extensions
                {
                    case ".fods":
                    case ".FODS":
                    case ".ods":
                    case ".ODS":
                    case ".ots":
                    case ".OTS":
                        // Inform user of input filepath
                        Console.WriteLine("---");
                        Console.WriteLine("VALIDATE");
                        Console.WriteLine(filepath);

                        // Validate file format standard
                        Validate_ODS ods = new Validate_ODS();
                        fileFormat_success = ods.Validate_Standard(filepath);

                        // Validate archival requirements
                        archivalReq_success = ods.Validate_ArchivalRequirements(filepath);

                        // Inform user of results
                        Console.WriteLine("---");
                        Console.WriteLine("SUMMARY");
                        if (fileFormat_success == false)
                        {
                            Console.WriteLine("--> File format standard: Invalid");
                        }
                        else if (fileFormat_success == true)
                        {
                            Console.WriteLine("--> File format standard: Valid");
                        }

                        if (archivalReq_success == false)
                        {
                            Console.WriteLine("--> Archival requirements: Invalid");
                        }
                        else if (archivalReq_success == true)
                        {
                            Console.WriteLine("--> Archival requirements: Valid");
                        }

                        // Return validation code
                        if (fileFormat_success == false || archivalReq_success == false)
                        {
                            return invalid;
                        }
                        else if (fileFormat_success == true && archivalReq_success == true)
                        {
                            return valid;
                        }
                        return fail;

                    case ".xlsb":
                    case ".XLSB":
                    case ".xlsm":
                    case ".XLSM":
                    case ".xlsx":
                    case ".XLSX":
                    case ".xltm":
                    case ".XLTM":
                    case ".xltx":
                    case ".XLTX":
                        // Inform user of input filepath
                        Console.WriteLine("---");
                        Console.WriteLine("VALIDATE");
                        Console.WriteLine(filepath);

                        // Validate file format standard
                        Validate_XLSX xlsx = new Validate_XLSX();
                        fileFormat_success = xlsx.Validate_Standard(filepath);

                        // Validate archival requirements
                        archivalReq_success = xlsx.Validate_ArchivalRequirements(filepath);

                        // Inform user of results
                        Console.WriteLine("---");
                        Console.WriteLine("SUMMARY");
                        if (fileFormat_success == false)
                        {
                            Console.WriteLine("--> File format standard: Invalid");
                        }
                        else if (fileFormat_success == true)
                        {
                            Console.WriteLine("--> File format standard: Valid");
                        }

                        if (archivalReq_success == false)
                        {
                            Console.WriteLine("--> Archival requirements: Invalid");
                        }
                        else if (archivalReq_success == true)
                        {
                            Console.WriteLine("--> Archival requirements: Valid");
                        }

                        // Return validation code
                        if (fileFormat_success == false || archivalReq_success == false)
                        {
                            Console.WriteLine("--> Spreadsheet is invalid");
                            return invalid;
                        }
                        else if (fileFormat_success == true && archivalReq_success == true)
                        {
                            Console.WriteLine("--> Spreadsheet is valid");
                            return valid;
                        }
                        return fail;

                    default:
                        Console.WriteLine("File format is not an accepted file format");
                        return fail;
                }
            }
            catch (FileNotFoundException) // If filepath has not file
            {
                Console.WriteLine("No file in filepath");
                return fail;
            }
            catch (FormatException) // If spreadsheet is password protected or otherwise unreadable
            {
                Console.WriteLine("File cannot be read");
                return fail;
            }
        }
    }
}