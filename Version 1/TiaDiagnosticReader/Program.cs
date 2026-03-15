using System;
using System.Windows.Forms;

namespace TiaDiagnosticGui
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Initializes the rendering of visual styles, fonts, etc.
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            // Runs the main form of the application
            Application.Run(new Form1());
        }
    }
}