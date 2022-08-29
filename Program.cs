using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

class Program
{
    public static string new_filepath = "";

    bool Main(string[] args)
    {
        string filepath = args[0];
        string extension = Path.GetExtension(filepath);
        bool success = true;
        bool fileFormat_success = true;
        bool archivalReq_success = true;

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
                    Console.WriteLine(filepath);
                    fileFormat_success = Validate_XLSX(filepath);
                    if (fileFormat_success == false)
                    {
                        success = false;
                    }
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
                    Console.WriteLine(filepath);
                    fileFormat_success = Validate_XLSX(filepath);
                    archivalReq_success = Validate_ArchivalRequirements_XLSX(filepath);
                    if (fileFormat_success == false || archivalReq_success == false)
                    {
                        success = false;
                    }
                    return success;
                default:
                    Console.WriteLine("File format is not an accepted file format");
                    success = false;
                    return success;
            }
        }
        catch (FileNotFoundException) // If filepath has not file
        {
            Console.WriteLine("No file in filepath");
            success = false;
            return success;
        }
        catch (FormatException) // If spreadsheet is password protected or otherwise unreadable
        {
            Console.WriteLine("File cannot be read");
            success = false;
            return success;
        }
    }

    static bool Validate_XLSX(string filepath)
    {
        bool success = true;
        using (var spreadsheet = SpreadsheetDocument.Open(filepath, false))
        {
            // Check for conformance
            bool? strict = spreadsheet.StrictRelationshipFound;
            if (strict == true)
            {
                Console.WriteLine($"--> File format is Strict conformant");
            }
            else
            {
                Console.WriteLine($"--> File format is Transitional conformant");
            }

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

    // Validate archival requirements (XLSX)
    public bool Validate_ArchivalRequirements_XLSX(string filepath)
    {
        bool success = true;

        bool data = Check_Value(filepath);
        if (data == true)
        {
            success = false;
        }
        int conn = Check_DataConnections(filepath);
        if (conn > 0)
        {
            success = false;
        }
        int cellrefs = Check_CellReferences(filepath);
        if (cellrefs > 0)
        {
            success = false;
        }
        int extobjs = Check_ExternalObjects(filepath);
        if (extobjs > 0)
        {
            success = false;
        }
        int rtdfunctions = Check_RTDFunctions(filepath);
        if (rtdfunctions > 0)
        {
            success = false;
        }
        int embedobjs = Check_EmbeddedObjects(filepath);
        if (embedobjs > 0)
        {
            success = false;
        }
        int printersettings = Check_PrinterSettings(filepath);
        if (printersettings > 0)
        {
            success = false;
        }
        bool activesheet = Check_ActiveSheet(filepath);
        if (activesheet == true)
        {
            success = false;
        }
        return success;
    }

    // Check for any values by checking if sheets and cell values exist
    public bool Check_Value(string filepath)
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
                    return hascellvalue;
                }
            }
        }
        Console.WriteLine("--> No cell values detected");
        return hascellvalue;
    }

    // Check for data connections
    public int Check_DataConnections(string filepath)
    {
        int conn_count = 0;

        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
        {
            ConnectionsPart conn = spreadsheet.WorkbookPart.ConnectionsPart;
            if (conn != null)
            {
                conn_count = conn.Connections.Count();
                Console.WriteLine($"--> {conn_count} data connections detected and removed");
            }
        }
        return conn_count;
    }

    // Check for external relationships
    public int Check_CellReferences(string filepath) // Find all external relationships
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
        if (cellreferences_count > 0)
        {
            Console.WriteLine($"--> {cellreferences_count} external cell references detected and removed");
        }
        return cellreferences_count;
    }

    // Check for external object references
    public int Check_ExternalObjects(string filepath)
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
        if (extobj_count > 0)
        {
            Console.WriteLine($"--> {extobj_count} external objects detected and removed");
        }
        return extobj_count;
    }

    // Check for RTD functions
    public static int Check_RTDFunctions(string filepath) // Check for RTD functions
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
        return rtd_functions_count;
    }

    // Check for embedded objects
    public int Check_EmbeddedObjects(string filepath) // Check for embedded objects and return alert
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
        return embedobj_count;
    }

    // Check for printer settings
    public int Check_PrinterSettings(string filepath)
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
            if (printerList.Count > 0)
            {
                Console.WriteLine($"--> {printersettings_count} printersettings detected and removed");
            }
        }
        return printersettings_count;
    }

    // Check for active sheet
    public bool Check_ActiveSheet(string filepath)
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
                    Console.WriteLine("--> First sheet is not active sheet detected and changed");
                    activeSheet = true;
                }
            }
        }
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

    public bool Validate_ArchivalRequirements_ODS(string filepath)
    {
        bool success = true;

        return success;
    }
}