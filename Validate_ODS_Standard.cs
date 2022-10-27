using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Validate.Spreadsheet
{
    public partial class Validate_ODS
    {
        public bool Validate_Standard(string filepath)
        {
            bool validity = false;

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
                    validity = false;
                }
                if (return_code == 1)
                {
                    Console.WriteLine("--> File format validation could not be completed");
                    validity = false;
                }
                if (return_code == 2)
                {
                    Console.WriteLine("--> File format is valid");
                    validity = true;
                }
                return validity;
            }
            catch (Win32Exception)
            {
                Console.WriteLine("--> File format validation requires ODF Validator and Java Development Kit");
                return validity;
            }
        }
    }
}
