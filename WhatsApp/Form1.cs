using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using System.Net.NetworkInformation;
using System.Data.SqlClient;
using Ini;

namespace ConsolaWhatsApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        [DllImport("user32.dll")]
        static extern bool ClientToScreen(IntPtr hWnd, ref Point lpPoint);

        [DllImport("user32.dll")]
        internal static extern uint SendInput(uint nInputs, [MarshalAs(UnmanagedType.LPArray), In] INPUT[] pInputs, int cbSize);

#pragma warning disable 649
        internal struct INPUT
        {
            public UInt32 Type;
            public MOUSEKEYBDHARDWAREINPUT Data;
        }

        [StructLayout(LayoutKind.Explicit)]
        internal struct MOUSEKEYBDHARDWAREINPUT
        {
            [FieldOffset(0)]
            public MOUSEINPUT Mouse;
        }

        internal struct MOUSEINPUT
        {
            public Int32 X;
            public Int32 Y;
            public UInt32 MouseData;
            public UInt32 Flags;
            public UInt32 Time;
            public IntPtr ExtraInfo;
        }



        [DllImport("user32.dll")]
        static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll")]
        static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll")]
        static extern bool MoveWindow(IntPtr Handle, int x, int y, int w, int h, bool repaint);

        static readonly int GWL_STYLE = -16;
        static readonly int WS_VISIBLE = 0x10000000;



        protected override void WndProc(ref Message message)
        {
            const int WM_SYSCOMMAND = 0x0112;
            const int SC_MOVE = 0xF010;

            switch (message.Msg)
            {
                case WM_SYSCOMMAND:
                    int command = message.WParam.ToInt32() & 0xfff0;
                    if (command == SC_MOVE)
                        return;
                    break;
            }

            base.WndProc(ref message);
        }

        public static PhysicalAddress GetMacAddress()
        {
            foreach (NetworkInterface nic in NetworkInterface.GetAllNetworkInterfaces())
            {
                // Only consider Ethernet network interfaces
                if (nic.NetworkInterfaceType == NetworkInterfaceType.Ethernet &&
                    nic.OperationalStatus == OperationalStatus.Up)
                {
                    return nic.GetPhysicalAddress();
                }
            }
            return null;
        }


        int Count = 0;
        int NGrupos;
        string[] Grupos = new string[1];
        string[] TIFFS = new string[1];
        DateTime[] LMDTS = new DateTime[1];

        string Startup = DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss");
        String TIFF, NombreArchivo, TXT, Archivo;
        DateTime LMDT;

        public string id = "0";
        public string Grupo, Mensaje;
        private void Form1_Load(object sender, EventArgs e)
        {
            //Process.Start(Environment.CurrentDirectory + @"\Resolution.exe");
            //Thread.Sleep(1000);


            try
            {
                this.MaximumSize = this.Size;
                this.MinimumSize = this.Size;
                panel0.Enabled = false;
                //TEVMANEBI01:005056904AF3; SSVMWSP001:0050568EE042; Notebook:28F10E0C5C1F
                if ((GetMacAddress().ToString() == "0050568EE042") || (GetMacAddress().ToString() == "28F10E0C5C1F") || true)
                {
                    IniFile ini = new IniFile(Environment.CurrentDirectory + @"\Setup.ini");
                    NGrupos = int.Parse(ini.IniReadValue("Grupos", "Cantidad"));
                    Array.Resize(ref Grupos, NGrupos);
                    Array.Resize(ref LMDTS, NGrupos);
                    Array.Resize(ref TIFFS, NGrupos);
                    int g;
                    for (int i = 0; i < Grupos.Length; i++)
                    {
                        g = i + 1;
                        Grupos[i] = ini.IniReadValue("Grupos", g.ToString());
                        LMDTS[i] = DateTime.Parse("01-01-2000");
                        TIFFS[i] = "-";
                    }




                    //Process p = Process.Start("whatsapp://");
                    Process p = Process.Start(Environment.CurrentDirectory + @"\WhatsApp\WhatsApp.lnk");
                    p.WaitForInputIdle();
                    while (p.MainWindowHandle == IntPtr.Zero)
                    {
                        Thread.Sleep(2500);
                        p.Refresh();
                    }

                    SetParent(p.MainWindowHandle, panel0.Handle);
                    SetWindowLong(p.MainWindowHandle, GWL_STYLE, WS_VISIBLE);
                    MoveWindow(p.MainWindowHandle, 0, 0, panel0.Width, panel0.Height, true);




                }
                else
                {
                    timer1.Interval = 500000000;
                    //MessageBox.Show(GetMacAddress().ToString());
                    button33.BackColor = Color.Red;
                }



                this.Text = "Inicio: " + Startup + " - " + "Total: " + Count.ToString();
                panel2.Focus();
                label1_1.Text = "...";
                label1_2.Text = label1_1.Text;
                label2_1.Text = label1_1.Text;
                label2_2.Text = label1_1.Text;
                label3_1.Text = label1_1.Text;
                label3_2.Text = label1_1.Text;
                label4_1.Text = label1_1.Text;
                label4_2.Text = label1_1.Text;
                label5_1.Text = label1_1.Text;
                label5_2.Text = label1_1.Text;
                label6_1.Text = label1_1.Text;
                label6_2.Text = label1_1.Text;
                label7_1.Text = label1_1.Text;
                label7_2.Text = label1_1.Text;
                label8_1.Text = label1_1.Text;
                label8_2.Text = label1_1.Text;
                label9_1.Text = label1_1.Text;
                label9_2.Text = label1_1.Text;
                label10_1.Text = label1_1.Text;
                label10_2.Text = label1_1.Text;
                label11_1.Text = label1_1.Text;
                label11_2.Text = label1_1.Text;
                label12_1.Text = label1_1.Text;
                label12_2.Text = label1_1.Text;
                label13_1.Text = label1_1.Text;
                label13_2.Text = label1_1.Text;
                label14_1.Text = label1_1.Text;
                label14_2.Text = label1_1.Text;
                label15_1.Text = label1_1.Text;
                label15_2.Text = label1_1.Text;
                label16_1.Text = label1_1.Text;
                label16_2.Text = label1_1.Text;
                label17_1.Text = label1_1.Text;
                label17_2.Text = label1_1.Text;
                label18_1.Text = label1_1.Text;
                label18_2.Text = label1_1.Text;
                label19_1.Text = label1_1.Text;
                label19_2.Text = label1_1.Text;
                label20_1.Text = label1_1.Text;
                label20_2.Text = label1_1.Text;

                //LastTIFF();
                LMDT = DateTime.Now; //System.IO.File.GetLastWriteTime(TIFF);

                timer1.Enabled = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Verifique:" + Environment.NewLine + "-El programa ya está en ejecución." + Environment.NewLine + "-No se encuentra Whatsapp.exe." + Environment.NewLine + "-" + ex.Message, "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1);
                this.Close();
            }

        }



        public static void ClickOnPoint(IntPtr wndHandle, Point clientPoint)
        {
            var oldPos = Cursor.Position;

            /// get screen coordinates
            ClientToScreen(wndHandle, ref clientPoint);

            /// set cursor on coords, and press mouse
            Cursor.Position = new Point(clientPoint.X, clientPoint.Y);

            var inputMouseDown = new INPUT();
            inputMouseDown.Type = 0; /// input type mouse
            inputMouseDown.Data.Mouse.Flags = 0x0002; /// left button down

            var inputMouseUp = new INPUT();
            inputMouseUp.Type = 0; /// input type mouse
            inputMouseUp.Data.Mouse.Flags = 0x0004; /// left button up

            var inputs = new INPUT[] { inputMouseDown, inputMouseUp };
            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));

            /// return mouse 
            Cursor.Position = oldPos;

        }


        #region
        #endregion


        int a = 1;
        Boolean led = true;
        int ContadorTIFF = 0;
        int dest;
        DateTime LMDT2 = DateTime.Parse("01-01-2000");
        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            if (led) { button33.BackColor = Color.LimeGreen; } else { button33.BackColor = Color.LightGray; }
            led = !led;

            if (ContadorTIFF > 0)
            {
                button33.BackColor = Color.Yellow;
                ContadorTIFF -= 1;
                if (ContadorTIFF < 0)
                    ContadorTIFF = 0;
            }
            else
            {


                try
                {

                    dest = -1;
                    int g;
                    string carpeta;
                    DateTime LMDTTIF, LMDTTXT;
                    for (int i = 0; i < Grupos.Length; i++)
                    {
                        if (Grupos[i] != "")
                        {

                            g = i + 1;
                            carpeta = "/Grupo" + g.ToString();
                            LastTIFF(carpeta);
                            LastTXT(carpeta);
                            LMDTTIF = System.IO.File.GetLastWriteTime(TIFF);
                            LMDTTXT = System.IO.File.GetLastWriteTime(TXT);

                            if (LMDTTIF > LMDTTXT)
                            {
                                LMDTS[i] = LMDTTIF;
                                TIFFS[i] = TIFF;
                            }
                            else
                            {
                                LMDTS[i] = LMDTTXT;
                                TIFFS[i] = TXT;
                            }


                            if (LMDTS[i] > LMDT2)
                            {
                                LMDT2 = LMDTS[i];
                                Archivo = TIFFS[i];
                                dest = i;
                            }
                        }
                    }

                    //this.Text = TIFF;

                    Boolean Es_Imagen = true;
                    if (Archivo.Contains(".txt")) { Es_Imagen = false; }


                    if ((LMDT2 > LMDT.AddSeconds(5)) & (dest > -1))
                    {

                        NombreArchivo = Path.GetFileNameWithoutExtension(Archivo);
                        if (this.WindowState != FormWindowState.Normal) { this.WindowState = FormWindowState.Normal; }

                        LMDT = LMDT2;
                        panel0.Enabled = true;

                        BuscarDestinatario(dest);

                        if (Es_Imagen)
                        {
                            Image Reporte;
                            Thread.Sleep(200);
                            using (FileStream stream = new FileStream(Archivo, FileMode.Open, FileAccess.Read))
                            {
                                Reporte = Image.FromStream(stream);
                                stream.Close();
                            }

                            Thread.Sleep(500);
                            Clipboard.SetImage(Reporte);
                            Thread.Sleep(500);
                            SendKeys.SendWait("^v");
                            Thread.Sleep(500);
                            try { Clipboard.SetText(NombreArchivo); }
                            catch { Clipboard.SetText("Reporte"); }
                            Thread.Sleep(500);
                            SendKeys.SendWait("^v");
                            Thread.Sleep(500);
                            SendKeys.SendWait("{ENTER}");
                            Thread.Sleep(500);
                        }
                        else
                        {
                            string MensajeTXT = File.ReadAllText(Archivo);

                            Thread.Sleep(500);
                            Clipboard.SetText(MensajeTXT);
                            Thread.Sleep(500);
                            SendKeys.SendWait("^v");
                            Thread.Sleep(500);
                            SendKeys.SendWait("{ENTER}");
                            Thread.Sleep(500);
                        }

                        Thread.Sleep(50);
                        Clipboard.Clear();
                        SendKeys.SendWait("^%/"); SendKeys.SendWait("^%/");

                        Thread.Sleep(50);

                        ContadorTIFF = 10;
                        Log(DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss"), NombreArchivo + " (" + (dest + 1).ToString() + ")");
                        panel0.Enabled = false;
                    }
                    else
                    {

                        SQL_GetMessage();

                        if (id != "0")
                        {
                            if (Int32.Parse(Grupo) <= NGrupos)
                            {
                                if (this.WindowState != FormWindowState.Normal) { this.WindowState = FormWindowState.Normal; }
                                dest = Int32.Parse(Grupo) - 1;

                                if (Grupos[dest] != "")
                                {
                                    panel0.Enabled = true;
                                    BuscarDestinatario(dest);

                                    Thread.Sleep(500);
                                    Clipboard.SetText(Mensaje);
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("^v");
                                    Thread.Sleep(500);
                                    SendKeys.SendWait("{ENTER}");
                                    Thread.Sleep(500);

                                    Clipboard.Clear();
                                    SendKeys.SendWait("^%/"); SendKeys.SendWait("^%/");
                                    Thread.Sleep(50);

                                    Log(DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss"), "SQL - " + id + " (" + (dest + 1).ToString() + ")");
                                    panel0.Enabled = false;
                                }
                            }

                            SQL_UPD_Message();
                        }
                    }

                    if (a > 15)
                    {
                        SQL_DT();
                        a = 1;
                    }
                    a++;

                }

                catch (Exception ex) { Log(DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss"), ex.Message); }
            }
            timer1.Enabled = true;
        }

        private void LastTIFF(string Carpeta)
        {
            var directory = new DirectoryInfo(Environment.CurrentDirectory + Carpeta);
            var LastModifiedFile = directory.GetFiles("*.tif").OrderByDescending(f => f.LastWriteTime).First();
            TIFF = LastModifiedFile.FullName;
        }

        private void LastTXT(string Carpeta)
        {
            var directory = new DirectoryInfo(Environment.CurrentDirectory + Carpeta);
            var LastModifiedFile = directory.GetFiles("*.txt").OrderByDescending(f => f.LastWriteTime).First();
            TXT = LastModifiedFile.FullName;
        }



        private void BuscarDestinatario(int idGrupo)
        {

            string DestWSP = Grupos[idGrupo];

            /*
            //ClickOnPoint(this.Handle, new Point(panel0.Location.X + 600, button1.Location.Y));
            //panel0.Enabled = true;
            ClickOnPoint(this.Handle, new Point(panel0.Location.X + 600, button5.Location.Y + 10));
            Thread.Sleep(500);
            //SendKeys.SendWait("^f");
            //Thread.Sleep(500);
            SendKeys.SendWait("^a");
            Thread.Sleep(500);
            SendKeys.SendWait(DestWSP);
            Thread.Sleep(500);
            SendKeys.SendWait("{ENTER}");
            Thread.Sleep(500);
            //panel0.Enabled = false;
            */


            //con control F funcionando. Sin control F, dejar el de arriba entre /**/
            /*
            ClickOnPoint(this.Handle, new Point(panel0.Location.X + 600, button1.Location.Y));
            panel0.Enabled = true;
            SendKeys.SendWait("^f");
            Thread.Sleep(500);
            SendKeys.SendWait("^a");
            Thread.Sleep(500);
            SendKeys.SendWait(DestWSP);
            Thread.Sleep(500);
            SendKeys.SendWait("{ENTER}");
            Thread.Sleep(500);
            */

            //con CTRL + ALT / (^%/) Para WhatsApp web. Reemplazar ^%/ por ^f para whatsapp windows   /**/
            ClickOnPoint(this.Handle, new Point(panel0.Location.X + 600, button1.Location.Y));
            panel0.Enabled = true;
            SendKeys.SendWait("^%/"); SendKeys.SendWait("^%/");
            Thread.Sleep(500);
            SendKeys.SendWait("^a");
            Thread.Sleep(500);
            SendKeys.SendWait(DestWSP);
            Thread.Sleep(500);
            SendKeys.SendWait("{ENTER}");
            Thread.Sleep(500);


        }

        private void Log(string DT, string Reporte)
        {

            //Application.Exit();

            Count += 1;
            this.Text = "Inicio: " + Startup + " - " + "Total: " + Count.ToString();

            label20_2.Text = label19_2.Text;
            label20_1.Text = label19_1.Text;
            label19_2.Text = label18_2.Text;
            label19_1.Text = label18_1.Text;
            label18_2.Text = label17_2.Text;
            label18_1.Text = label17_1.Text;
            label17_2.Text = label16_2.Text;
            label17_1.Text = label16_1.Text;
            label16_2.Text = label15_2.Text;
            label16_1.Text = label15_1.Text;
            label15_2.Text = label14_2.Text;
            label15_1.Text = label14_1.Text;
            label14_2.Text = label13_2.Text;
            label14_1.Text = label13_1.Text;
            label13_2.Text = label12_2.Text;
            label13_1.Text = label12_1.Text;
            label12_2.Text = label11_2.Text;
            label12_1.Text = label11_1.Text;
            label11_2.Text = label10_2.Text;
            label11_1.Text = label10_1.Text;
            label10_2.Text = label9_2.Text;
            label10_1.Text = label9_1.Text;
            label9_2.Text = label8_2.Text;
            label9_1.Text = label8_1.Text;
            label8_2.Text = label7_2.Text;
            label8_1.Text = label7_1.Text;
            label7_2.Text = label6_2.Text;
            label7_1.Text = label6_1.Text;
            label6_2.Text = label5_2.Text;
            label6_1.Text = label5_1.Text;
            label5_2.Text = label4_2.Text;
            label5_1.Text = label4_1.Text;
            label4_2.Text = label3_2.Text;
            label4_1.Text = label3_1.Text;
            label3_2.Text = label2_2.Text;
            label3_1.Text = label2_1.Text;
            label2_2.Text = label1_2.Text;
            label2_1.Text = label1_1.Text;

            label1_2.Text = Reporte;
            label1_1.Text = DT;


        }


        string connectionString = "Server=localhost;Database=WhatsApp;User Id=sa;Password=H0n3ywell!;";
        private void SQL_DT()
        {
            string queryString = "UPDATE [dbo].[Monitoreo] SET [DT] = GETDATE()";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(queryString, connection);
                connection.Open();
                command.ExecuteNonQuery();
            }
        }


        private void SQL_GetMessage()
        {

            string queryString = "SELECT Top 1 cast(id as varchar) as id,cast(Grupo as varchar) as Grupo,Mensaje FROM [dbo].[Mensajes] WHERE Estado=0 AND DT>GETDATE()-0.125 and Grupo in (7,8,10,11,12,13,14,18,19,21,22,31,32) ORDER BY id";
            id = "0";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(queryString, connection);
                connection.Open();
                SqlDataReader reader = command.ExecuteReader();
                try
                {
                    while (reader.Read())
                    {
                        id = reader["id"].ToString();
                        Grupo = reader["Grupo"].ToString();
                        Mensaje = reader["Mensaje"].ToString();
                    }
                }
                finally
                {
                    reader.Close();
                }
            }
        }

        private void SQL_UPD_Message()
        {
            string queryString = "UPDATE [dbo].[Mensajes] SET [Estado] = 1 WHERE id=" + id;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(queryString, connection);
                connection.Open();
                command.ExecuteNonQuery();
            }
        }


    }
}
