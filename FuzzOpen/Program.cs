using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace FuzzOpen
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
           // AppDomain currentDomain = AppDomain.CurrentDomain;
           // currentDomain.UnhandledException += currentDomain_UnhandledException;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
        
        
        static void currentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            MessageBox.Show("This is shown when ANY thread is thrown in ANY point of your Domain.");
        }
    }
}