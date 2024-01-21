// System
using System.Drawing;
using System.Diagnostics;
using System.Runtime.InteropServices;

// Google Cloud Translation API
using Google.Cloud.Translation.V2;

// Add the following assembly references to the project for PowerPoint automation
//
// C:\Program Files (x86)\Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15\Office.dll
// C:\Program Files (x86)\Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15\Microsoft.Office.Interop.Excel.dll
// C:\Program Files (x86)\Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15\Microsoft.Office.Interop.PowerPoint.dll
//
// Repair Office 365 (Online mode) in case of any COM errors
//
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPtrans
{
    internal class Program
    {
        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly, string lpWindowName);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr WindowHandle);

        static void Main(string[] args)
        {
            try
            {
                // Run in a safe context
                MainSafe(args);
            }
            catch (Exception e)
            {
                // Show error message
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);

            }
            finally
            {
                // Kill orphaned processes
                Kill("EXCEL", "");
                Kill("POWERPNT", "PowerPoint");
                Console.ReadLine();
            }
        }

        /// <summary>
        /// Defines the entry point of the application (Safe Context).
        /// </summary>
        /// <param name="args">The arguments.</param>
        static void MainSafe(string[] args)
        {
            // Check command line arguments
            if (args.Length < 1) throw new Exception("File name missing!");
            if (args.Length < 2) throw new Exception("Language!");

            // Pickup filename and language
            string file = Path.GetFullPath(args[0]);
            string lang = args[1];

            // Detect file type
            if (file.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                TranslateExcel(file, lang);
            }
            else if (file.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
            {
                TranslatePowerPoint(file, lang);
            }
        }

        /// <summary>
        /// Translates the Excel workbook.
        /// </summary>
        static void TranslateExcel(string file, string lang)
        {
            // Translate text
            string txt = null;

            // Google translation client
            var gtc = TranslationClient.Create();

            // Launch Excel
            var exa = new Excel.Application();

            // Open workbook
            Workbook wb = exa.Workbooks.Open(file, ReadOnly: true);

            // Iterate all worksheets
            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                // Action counter
                int i = 0;

                // Onlide slide worksheets
                if (ws.Name.StartsWith("Slide", StringComparison.OrdinalIgnoreCase))
                {
                    while (ws.Range["B2"].Offset[i, 0].Text.ToString() != "")
                    {
                        // Get argument
                        string arg = ws.Range["B2"].Offset[i, 0].Text.ToString();

                        // Check type
                        bool ctrl = ("#@~".Contains(arg[0]));

                        // Only translate non-control arguments
                        if (!ctrl)
                        {
                            // Make the API call
                            txt = gtc.TranslateText(arg, lang, "en").TranslatedText;

                            // Update the workbook
                            ws.Range["B2"].Offset[i, 0].Value = txt;

                            // Show progress
                            Console.WriteLine(arg);
                            Console.WriteLine(txt);
                            Console.WriteLine();
                        }

                        // Increment action counter
                        i++;
                    }
                }
            }

            // Update the file with the language code
            file = Path.Combine(Path.GetDirectoryName(file), Path.GetFileNameWithoutExtension(file) + "." + lang + ".xlsx");

            // Save the workbook
            wb.SaveAs(file);

            // Progress 
            Console.WriteLine("Output: " + file);
            Console.WriteLine();
            Console.WriteLine("Ready. Please check workbook for translation errors!");

            // Quit Excel
            try { wb.Close(); } catch { }
            try { exa.Quit(); } catch { }
            try { Kill("EXCEL", ""); } catch { }
        }

        /// <summary>
        /// Translates the Excel workbook.
        /// </summary>
        static void TranslatePowerPoint(string file, string lang)
        {
            //string[] fonts = { "Consolas", "Courier New", "Lucida Sans Typewriter" };

            // Translate text
            string txt = null;

            // Google translation client
            var gtc = TranslationClient.Create();

            // Launch powerpoint
            var ppa = new PowerPoint.Application();

            // Open presentation
            var ppp = ppa.Presentations.Open(file);

            // Move the focus to the console window
            FocusWindow(Console.Title);

            // Iterate all slides
            foreach (PowerPoint.Slide slide in ppp.Slides)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        // Analyze font
                        string fname  = shape.TextFrame.TextRange.Font.Name;
                        float fsize   = shape.TextFrame.TextRange.Font.Size;
                        var font      = new System.Drawing.Font(fname, fsize);
                        var monospace = IsMonospacedFont(font);
                        var equation  = fname.Contains("Math", StringComparison.OrdinalIgnoreCase);
                        var exclude   = (shape.AlternativeText == "@");

                        // We assume monospaced fonts represent code that does not require translation
                        // We also do not want to touch equations because all formats will be lost
                        // Set the shape Alt Text to "@" to exclude the shape from translation

                        if (!monospace && !equation && !exclude)
                        {
                            // Preserve layout by sizing the text to the shape
                            shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;

                            // Get the text
                            string arg = shape.TextFrame2.TextRange.Text.Trim();

                            // Check for empty strings
                            if (!string.IsNullOrEmpty(arg))
                            {
                                // Make the API call
                                txt = gtc.TranslateText(arg, lang, "en").TranslatedText;

                                // Update the presentation
                                shape.TextFrame2.TextRange.Text = txt;                                  // This destroys all formatting
                                //shape.TextFrame.TextRange.Font.Color.RGB = Color.Black.ToArgb();      // Reset color to black

                                // Show progress
                                Console.WriteLine(arg);
                                Console.WriteLine(txt);
                                Console.WriteLine();
                            }
                        }
                    }
                }
            }

            // Update the file with the language code
            file = Path.Combine(Path.GetDirectoryName(file), Path.GetFileNameWithoutExtension(file) + "." + lang + ".pptx");

            // Save the workbook
            ppp.SaveAs(file);

            // Progress 
            Console.WriteLine("Output: " + file);
            Console.WriteLine();
            Console.WriteLine("Ready. Please check presentation for colors, layout and translation errors!");

            // Quit PowerPoint          
            try { ppp.Close(); } catch { }
            try { ppa.Quit(); } catch { }
            try { Kill("POWERPNT", "PowerPoint"); } catch { }
        }
        public static bool IsMonospacedFont(System.Drawing.Font font)
        {
            using (Bitmap bitmap = new Bitmap(1, 1))
            {
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    // Choose two characters with different widths, e.g., 'i' and 'W'
                    string testString = "iW";
                    SizeF size1 = g.MeasureString(testString[0].ToString(), font);
                    SizeF size2 = g.MeasureString(testString[1].ToString(), font);

                    // If the widths are different, it's not a monospaced font
                    return size1.Width == size2.Width;
                }
            }
        }

        /// <summary>Moves the focus to window.</summary>
        static void FocusWindow(string title)
        {
            FocusWindow(FindWindowByCaption(0, Console.Title));
        }

        /// <summary>Moves the focus to window.</summary>
        static void FocusWindow(IntPtr hwnd)
        {
            SetForegroundWindow(hwnd);
        }

        /// <summary>
        /// Kills the process indentified by its signature.
        /// </summary>
        /// <param name="signature"></param>
        static void Kill(string signature, string title)
        {
            // Create an array of all running Excel processes
            Process[] processes = Process.GetProcessesByName(signature);

            // Loop over these processes
            foreach (var process in processes)
            {
                // Only look at the instance with an empty window title
                if (process.MainWindowTitle == title)
                {
                    // Kill the process
                    process.Kill();
                }
            }
        }
    }
}
