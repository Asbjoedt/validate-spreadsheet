using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace Validate.Spreadsheet
{
    public partial class Validate_XLSX
    {
        public bool Validate_Standard(string filepath)
        {
            bool success = false;

            Console.WriteLine("---");
            Console.WriteLine("FILE FORMAT STANDARD");

            using (var spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                // Validate
                var validator = new OpenXmlValidator();
                var validation_errors = validator.Validate(spreadsheet).ToList();
                int error_count = validation_errors.Count;
                int error_number = 0;
                if (validation_errors.Any()) // If errors, inform user & return results
                {
                    Console.WriteLine($"--> File format has {error_count} validation errors");
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
                    return success;
                }
                else
                {
                    Console.WriteLine("--> Spreadsheet is compliant with file format standard");
                    success = true;
                    return success;
                }
            }
        }
    }
}
