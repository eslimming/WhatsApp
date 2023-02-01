using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace ConsolaWhatsApp
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());

            bool instanceCountOne = false;


            using (System.Threading.Mutex mtex = new System.Threading.Mutex(true, "Consola_WhatsApp_EFSQ", out instanceCountOne))
            {
                if (instanceCountOne)
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new Form1());
                    mtex.ReleaseMutex();
                }
                else
                {
                    //MessageBox.Show("An application instance is already running!!!!!!!");
                }
            }

        }
    }
}
