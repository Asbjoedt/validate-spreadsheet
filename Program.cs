using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

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
                    fileFormat_success = Validate_Standard_ODS(filepath);

                    // Validate archival data quality specifications
                    archivalReq_success = Validate_ArchivalRequirements_ODS(filepath);

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
                    fileFormat_success = Validate_Standard_XLSX(filepath);

                    // Validate archival data quality specifications
                    archivalReq_success = Validate_ArchivalRequirements_XLSX(filepath);

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

    static bool Validate_Standard_XLSX(string filepath)
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

    // Validate archival requirements (XLSX)
    static bool Validate_ArchivalRequirements_XLSX(string filepath)
    {
        bool success = false;

        Console.WriteLine("---");
        Console.WriteLine("ARCHIVAL REQUIREMENTS");

        bool strict = Check_Strict(filepath);

        bool data = Check_Value(filepath);

        int conn = Check_DataConnections(filepath);

        int cellrefs = Check_CellReferences(filepath);

        int extobjs = Check_ExternalObjects(filepath);

        int rtdfunctions = Check_RTDFunctions(filepath);

        int embedobjs = Check_EmbeddedObjects(filepath);

        int printersettings = Check_PrinterSettings(filepath);

        bool activesheet = Check_ActiveSheet(filepath);

        if (strict == true && data == true && conn == 0 && cellrefs == 0 && extobjs == 0 && rtdfunctions == 0 && embedobjs == 0 && printersettings == 0 && activesheet == false)
        {
            success = true;
            return success;
        }
        else
        {
            success = false;
            return success;
        }
    }

    // Check for Strict conformance
    static bool Check_Strict(string filepath)
    {
        bool strict = false;

        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
        {
            Workbook workbook = spreadsheet.WorkbookPart.Workbook;
            if (workbook.Conformance == "strict")
            {
                Console.WriteLine("--> Spreadsheet is Strict conformant");
                strict = true;
            }
            else if (workbook.Conformance == null || workbook.Conformance == "transitional")
            {
                Console.WriteLine("--> Spreadsheet is Transitional conformant (Error)");
                strict = false;
            }
        }
        return strict;
    }

    // Check for any values by checking if sheets and cell values exist
    static bool Check_Value(string filepath)
    {
        bool hascellvalue = false;

        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
        {
            //Check if worksheets exist
            WorkbookPart wbPart = spreadsheet.WorkbookPart;
            DocumentFormat.OpenXml.Spreadsheet.Sheets allSheets = wbPart.Workbook.Sheets;
            if (allSheets == null)
            {
                Console.WriteLine("--> No cell values detected");
                return hascellvalue;
            }
            // Check if any cells have any value
            foreach (Sheet aSheet in allSheets)
            {
                WorksheetPart wsp = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(aSheet.Id);
                Worksheet worksheet = wsp.Worksheet;
                var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>(); // Find all rows
                int row_count = rows.Count(); // Count number of rows
                if (row_count > 0) // If any rows exist, this means cells exist
                {
                    hascellvalue = true;
                    Console.WriteLine("--> Cell values detected");
                    return hascellvalue;
                }
            }
        }
        Console.WriteLine("--> No cell values detected");
        return hascellvalue;
    }

    // Check for data connections
    static int Check_DataConnections(string filepath)
    {
        int conn_count = 0;

        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
        {
            ConnectionsPart conn = spreadsheet.WorkbookPart.ConnectionsPart;
            if (conn != null)
            {
                conn_count = conn.Connections.Count();
            }
        }
        Console.WriteLine($"--> {conn_count} data connections detected");
        return conn_count;
    }

    // Check for external relationships
    static int Check_CellReferences(string filepath) // Find all external relationships
    {
        int cellreferences_count = 0;

        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
        {
            List<ExternalWorkbookPart> extwbParts = spreadsheet.WorkbookPart.ExternalWorkbookParts.ToList();
            if (extwbParts.Count > 0)
            {
                foreach (ExternalWorkbookPart ext in extwbParts)
                {
                    var elements = ext.ExternalLink.ChildElements.ToList();
                    foreach (var element in elements)
                    {
                        if (element.LocalName == "externalBook")
                        {
                            var externalLink = ext.ExternalLink.ToList();
                            foreach (ExternalBook externalBook in externalLink)
                            {
                                var cellreferences = externalBook.SheetDataSet.ChildElements.ToList();
                                foreach (var cellreference in cellreferences)
                                {
                                    var cells = cellreference.InnerText.ToList();
                                    foreach (var cell in cells)
                                    {
                                        cellreferences_count++;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        Console.WriteLine($"--> {cellreferences_count} external cell references detected");
        return cellreferences_count;
    }

    // Check for external object references
    static int Check_ExternalObjects(string filepath)
    {
        int extobj_count = 0;

        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
        {
            List<ExternalWorkbookPart> extwbParts = spreadsheet.WorkbookPart.ExternalWorkbookParts.ToList();
            if (extwbParts.Count > 0)
            {
                foreach (ExternalWorkbookPart ext in extwbParts)
                {
                    var elements = ext.ExternalLink.ChildElements.ToList();
                    foreach (var element in elements)
                    {
                        if (element.LocalName == "oleLink")
                        {
                            var externalLink = ext.ExternalLink.ToList();
                            foreach (OleLink oleLink in externalLink)
                            {
                                extobj_count++;
                            }
                        }
                    }
                }
            }
        }
        Console.WriteLine($"--> {extobj_count} external objects detected");
        return extobj_count;
    }

    // Check for RTD functions
    static int Check_RTDFunctions(string filepath) // Check for RTD functions
    {
        int rtd_functions_count = 0;

        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
        {
            WorkbookPart wbPart = spreadsheet.WorkbookPart;
            Sheets allSheets = wbPart.Workbook.Sheets;
            foreach (Sheet aSheet in allSheets)
            {
                WorksheetPart wsp = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(aSheet.Id);
                Worksheet worksheet = wsp.Worksheet;
                var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>(); // Find all rows
                foreach (var row in rows)
                {
                    var cells = row.Elements<Cell>();
                    foreach (Cell cell in cells)
                    {
                        if (cell.CellFormula != null)
                        {
                            string formula = cell.CellFormula.InnerText;
                            if (formula.Length > 2)
                            {
                                string hit = formula.Substring(0, 3); // Transfer first 3 characters to string
                                if (hit == "RTD")
                                {
                                    rtd_functions_count++;
                                    Console.WriteLine($"--> RTD function in sheet \"{aSheet.Name}\" cell {cell.CellReference} detected and removed");
                                }
                            }
                        }
                    }
                }
            }
        }
        Console.WriteLine($"--> {rtd_functions_count} RTD functions detected");
        return rtd_functions_count;
    }

    // Check for embedded objects
    static int Check_EmbeddedObjects(string filepath) // Check for embedded objects and return alert
    {
        int embedobj_count = 0;

        using (var spreadsheet = SpreadsheetDocument.Open(filepath, false))
        {
            var list = spreadsheet.WorkbookPart.WorksheetParts.ToList();
            foreach (var item in list)
            {
                int count_ole = item.EmbeddedObjectParts.Count(); // Register the number of OLE
                int count_image = item.ImageParts.Count(); // Register number of images
                int count_3d = item.Model3DReferenceRelationshipParts.Count(); // Register number of 3D models
                embedobj_count = count_ole + count_image + count_3d; // Sum

                if (embedobj_count > 0) // If embedded objects
                {
                    Console.WriteLine($"--> {embedobj_count} embedded objects detected");
                    var embed_ole = item.EmbeddedObjectParts.ToList(); // Register each OLE to a list
                    var embed_image = item.ImageParts.ToList(); // Register each image to a list
                    var embed_3d = item.Model3DReferenceRelationshipParts.ToList(); // Register each 3D model to a list
                    int embedobj_number = 0;
                    foreach (var part in embed_ole) // Inform user of each OLE object
                    {
                        embedobj_number++;
                        Console.WriteLine($"--> Embedded object #{embedobj_number}");
                        Console.WriteLine($"----> Content Type: {part.ContentType.ToString()}");
                        Console.WriteLine($"----> URI: {part.Uri.ToString()}");

                    }
                    foreach (var part in embed_image) // Inform user of each image object
                    {
                        embedobj_number++;
                        Console.WriteLine($"--> Embedded object #{embedobj_number}");
                        Console.WriteLine($"----> Content Type: {part.ContentType.ToString()}");
                        Console.WriteLine($"----> URI: {part.Uri.ToString()}");
                    }
                    foreach (var part in embed_3d) // Inform user of each 3D object
                    {
                        embedobj_number++;
                        Console.WriteLine($"--> Embedded object #{embedobj_number}");
                        Console.WriteLine($"----> Content Type: {part.ContentType.ToString()}");
                        Console.WriteLine($"----> URI: {part.Uri.ToString()}");
                    }
                }
            }
        }
        Console.WriteLine($"--> {embedobj_count} embedded objects detected");
        return embedobj_count;
    }

    // Check for printer settings
    static int Check_PrinterSettings(string filepath)
    {
        int printersettings_count = 0;

        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
        {
            var worksheetpartslist = spreadsheet.WorkbookPart.WorksheetParts.ToList();
            List<SpreadsheetPrinterSettingsPart> printerList = new List<SpreadsheetPrinterSettingsPart>();
            foreach (WorksheetPart worksheetpart in worksheetpartslist)
            {
                printerList = worksheetpart.SpreadsheetPrinterSettingsParts.ToList();
            }
            foreach (SpreadsheetPrinterSettingsPart printer in printerList)
            {
                printersettings_count++;
            }
        }
        Console.WriteLine($"--> {printersettings_count} printersettings detected");
        return printersettings_count;
    }

    // Check for active sheet
    static bool Check_ActiveSheet(string filepath)
    {
        bool activeSheet = false;

        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
        {
            BookViews bookViews = spreadsheet.WorkbookPart.Workbook.GetFirstChild<BookViews>();
            WorkbookView workbookView = bookViews.GetFirstChild<WorkbookView>();
            if (workbookView.ActiveTab != null)
            {
                var activeSheetId = workbookView.ActiveTab.Value;
                if (activeSheetId > 0)
                {
                    Console.WriteLine("--> First sheet is not active sheet detected");
                    activeSheet = true;
                    return activeSheet;
                }
            }
        }
        Console.WriteLine("--> First sheet is active sheet detected");
        return activeSheet;
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

    static bool Validate_Standard_ODS(string filepath)
    {
        bool valid = false;

        try
        {
            // Use ODF Validator for validation of OpenDocument spreadsheets
            Process app = new Process();
            app.StartInfo.UseShellExecute = false;
            app.StartInfo.FileName = "javaw";
            string normal_dir = "C:\\Program Files\\ODF Validator\\odfvalidator-0.10.0-jar-with-dependencies.jar";
            string? environ_dir = null;
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) // If app is run on Windows
            {
                environ_dir = Environment.GetEnvironmentVariable("ODFValidator");
            }
            if (environ_dir != null)
            {
                app.StartInfo.Arguments = $"-jar \"{environ_dir}\" \"{filepath}\"";
            }
            else
            {
                app.StartInfo.Arguments = $"-jar \"{normal_dir}\" \"{filepath}\"";
            }
            app.Start();
            app.WaitForExit();
            int return_code = app.ExitCode;
            app.Close();

            // Inform user of validation results
            if (return_code == 0)
            {
                Console.WriteLine("--> File format is invalid. Spreadsheet has no cell values");
                valid = false;
            }
            if (return_code == 1)
            {
                Console.WriteLine("--> File format validation could not be completed");
                valid = false;
            }
            if (return_code == 2)
            {
                Console.WriteLine("--> File format is valid");
                valid = true;
            }
            return valid;
        }
        catch (Win32Exception)
        {
            Console.WriteLine("--> File format validation requires ODF Validator and Java Development Kit");
            return valid;
        }
    }

    static bool Validate_ArchivalRequirements_ODS(string filepath)
    {
        bool success = true;

        return success;
    }
}