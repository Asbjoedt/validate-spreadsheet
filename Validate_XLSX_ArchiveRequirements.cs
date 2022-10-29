using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Validate.Spreadsheet
{
    public partial class Validate_XLSX
    {
        public bool Validate_ArchivalRequirements(string filepath)
        {
            bool success = false;

            // Inform user
            Console.WriteLine("Validating archival requirements");

            // Perform checks
            bool strict = Check_Conformance(filepath);
            bool data = Check_Value(filepath);
            int conn = Check_DataConnections(filepath);
            int cellrefs = Check_CellReferences(filepath);
            int extobjs = Check_ExternalObjects(filepath);
            int rtdfunctions = Check_RTDFunctions(filepath);
            int printersettings = Check_PrinterSettings(filepath);
            bool activesheet = Check_ActiveSheet(filepath);
            //int embedobjs = Check_EmbeddedObjects(filepath);

            // Return success
            if (strict == true && data == true && conn == 0 && cellrefs == 0 && extobjs == 0 && rtdfunctions == 0 && printersettings == 0 && activesheet == false)
            {
                Console.WriteLine("Archival requirements: Valid");
                success = true;
                return success;
            }
            else
            {
                Console.WriteLine("Archival requirements: Invalid");
                success = false;
                return success;
            }
        }

        // Check for Strict conformance
        static bool Check_Conformance(string filepath)
        {
            bool strict = false;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath, false))
            {
                Workbook workbook = spreadsheet.WorkbookPart.Workbook;
                if (workbook.Conformance == null || workbook.Conformance.Value == ConformanceClass.Enumtransitional)
                {
                    Console.WriteLine("Error: Transitional conformance detected");
                    strict = false;
                }
                else if (workbook.Conformance.Value == ConformanceClass.Enumstrict)
                {
                    strict = true;
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
                Sheets allSheets = wbPart.Workbook.Sheets;
                if (allSheets == null)
                {
                    Console.WriteLine("Error: No cell values detected");
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
            Console.WriteLine("Error: No cell values detected");
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
                    foreach (Connection c in conn.Connections)
                    {
                        Console.WriteLine($"Error: Data connection \"{c.NamespaceUri}\" detected");
                    }
                }
            }
            Console.WriteLine($"Error: In total {conn_count} data connections detected");
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
                                            Console.WriteLine($"Error: External cell reference detected");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            Console.WriteLine($"Error: In total {cellreferences_count} external cell references detected");
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
                                    Console.WriteLine($"Error: External object \"{oleLink.NamespaceUri}\" detected");
                                }
                            }
                        }
                    }
                }
            }
            Console.WriteLine($"Error: In total {extobj_count} external objects detected");
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
                                        Console.WriteLine($"Error: RTD function in sheet \"{aSheet.Name}\" cell {cell.CellReference} detected and removed");
                                    }
                                }
                            }
                        }
                    }
                }
            }
            Console.WriteLine($"Error: In total {rtd_functions_count} RTD functions detected");
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
                        Console.WriteLine($"{embedobj_count} embedded objects detected");
                        var embed_ole = item.EmbeddedObjectParts.ToList(); // Register each OLE to a list
                        var embed_image = item.ImageParts.ToList(); // Register each image to a list
                        var embed_3d = item.Model3DReferenceRelationshipParts.ToList(); // Register each 3D model to a list
                        int embedobj_number = 0;
                        foreach (var part in embed_ole) // Inform user of each OLE object
                        {
                            embedobj_number++;
                            Console.WriteLine($"Embedded object #{embedobj_number}");
                            Console.WriteLine($"--> Content Type: {part.ContentType.ToString()}");
                            Console.WriteLine($"--> URI: {part.Uri.ToString()}");

                        }
                        foreach (var part in embed_image) // Inform user of each image object
                        {
                            embedobj_number++;
                            Console.WriteLine($"Embedded object #{embedobj_number}");
                            Console.WriteLine($"--> Content Type: {part.ContentType.ToString()}");
                            Console.WriteLine($"--> URI: {part.Uri.ToString()}");
                        }
                        foreach (var part in embed_3d) // Inform user of each 3D object
                        {
                            embedobj_number++;
                            Console.WriteLine($"Embedded object #{embedobj_number}");
                            Console.WriteLine($"--> Content Type: {part.ContentType.ToString()}");
                            Console.WriteLine($"--> URI: {part.Uri.ToString()}");
                        }
                    }
                }
            }
            Console.WriteLine($"{embedobj_count} embedded objects detected");
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
                    Console.WriteLine("Error: Printer setting detected");
                    printersettings_count++;
                }
            }
            Console.WriteLine($"Error: In total {printersettings_count} printersettings detected");
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
                        Console.WriteLine("Error: First sheet is not active sheet detected");
                        activeSheet = true;
                        return activeSheet;
                    }
                }
            }
            Console.WriteLine("First sheet is active sheet detected");
            return activeSheet;
        }
    }
}
