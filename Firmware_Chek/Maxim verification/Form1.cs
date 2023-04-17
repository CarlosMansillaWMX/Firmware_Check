using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.IO;
using System.Runtime.InteropServices;

namespace Maxim_verification
{
    public partial class Form1 : Form
    {
        string diadelasemana = "";
        string dianum = "";
        string mes = "";
        string hora = "";
        string minutos = "";
        Aspose.Cells.Workbook wb1 = new Aspose.Cells.Workbook();
        private String filePath = Directory.GetCurrentDirectory() + "\\Config.ini";
        //SFC_System wmx_ts = new SFC_System();
        Form f2;
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetDesktopWindow();
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetWindowDC(IntPtr window);
        [DllImport("gdi32.dll", SetLastError = true)]
        public static extern uint GetPixel(IntPtr dc, int x, int y);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern int ReleaseDC(IntPtr window, IntPtr dc);
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);
        public void Write(string section, string key, string value)
        {
            WritePrivateProfileString(section, key, value, this.filePath);
        }
        public string Read(string section, string key)
        {
            StringBuilder SB = new StringBuilder(255);
            int i = GetPrivateProfileString(section, key, "", SB, 255, this.filePath);
            return SB.ToString();
        }
        public Form1(Form2 f2)
        {

            InitializeComponent();
            this.f2 = f2;
            timer1.Enabled = true;
            wb1.Worksheets.Add("Log");
            if (Read("BIRD1", "Port") == String.Empty)
            {
                Write("BIRD1", "Port", "COM5");
            }
            serialPort1.PortName = Read("BIRD1", "Port");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            string Teststart = DateTime.Now.ToString("MM/dd/yyyy HH\\:mm\\:ss");
            tbcom.Text = "";
            panel1.BackColor = Color.SteelBlue;
            string[] subs = new string[6] { "", "", "", "", "", "" };
            tbcom.Text = "";
            string loco = "";
            string SN =tbsn.Text;
            string revision = "";
            string VV1 = "";
            int conect = 0;
            string[] subs2 = new string[6] { "", "", "", "", "", "" };
            int i = 0;
            int t = 0;
            Stopwatch sw1 = new Stopwatch(); 
            
            sw1.Start(); 
            if (SN.Length >= 9 && SN.Length > 0)
            {
                tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + "Check NFC Conection" + $"{Environment.NewLine}" + "X05,01,24,4,68A3" + $"{Environment.NewLine}" + "W01,sync";
                serialPort1.Open();
                while (true)
                {
                    serialPort1.DiscardOutBuffer();
                    serialPort1.DiscardInBuffer();  
                    serialPort1.WriteLine("X05,01,24,4,68A3");
                    try
                    {
                        VV1 = "";
                        VV1 = serialPort1.ReadLine();
                    }
                    catch (Exception)
                    {
                        conect = conect + 1;
                        tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + "no response";
                    }
                    tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + VV1;
                    serialPort1.DiscardOutBuffer();
                    serialPort1.DiscardInBuffer();
                    serialPort1.WriteLine("W01,sync");
                    try
                    {
                        VV1 = "";
                        VV1 = serialPort1.ReadLine();
                    }
                    catch (Exception)
                    {
                        conect = conect + 1;
                    }
                    tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + VV1;
                    if ( VV1== "W02,ack")
                    {
                        i = 7;
                        conect = 0;
                        break;
                        
                    }
                    else
                    {
                        Thread.Sleep(1000);
                        i = i + 1;
                        tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + "\r" + i.ToString();
                    }
                    if (i == 10)
                    {
                        conect = 1;
                        break;
                    }
                }
                if (conect == 0)
                {
                    tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + "Unlock Quiet mode" + $"{Environment.NewLine}" + "m085 2";
                    //adicion de nuevo renglon
                    int n = dataGridView1.Rows.Add();

                    //colocacion de informacion 
                    dataGridView1.Rows[n].Cells[0].Value = "NFC Conection";
                    dataGridView1.Rows[n].Cells[1].Value = "Pass";
                    dataGridView1.Rows[n].Cells[2].Value = "Pass";
                    dataGridView1.Rows[n].Cells[3].Value = "Pass";
                    dataGridView1.Rows[n].Cells[4].Value = "Pass";
                    dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Green;
                    i = 0;
                    while (true)
                    {
                        if (VV1 != "")
                        {

                            tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + VV1;
                            serialPort1.DiscardOutBuffer();
                            serialPort1.DiscardInBuffer();
                            serialPort1.WriteLine("m085 2");
                            try
                            {
                                VV1 = serialPort1.ReadLine();
                            }
                            catch (Exception){}
                            tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + VV1;
                        }
                        if (VV1.Length > 1)
                            subs = VV1.Split(',');
                        if (subs.Length > 2)
                        {
                            tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + subs[2];
                            if (subs[2].Length >= 12)
                            {
                                tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + subs[2].Substring(0, 12);
                                loco = subs[2].Substring(0, 12);
                                tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + loco;
                            }
                        }
                        if (loco == ":m085 002 OK")
                        {
                            i = 7;
                            break;
                        }
                        else
                        {
                            i++;
                            Thread.Sleep(1000);
                            tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + "\r" + i.ToString();
                            VV1="Attemp"+i.ToString();
                        }
                        if(i==15)
                        {
                            break;
                        }
                    }

                    i = 0;
                    VV1 = "";
                    while (true)
                    {
                        subs = new string[6] { "", "", "", "", "", "" };
                        subs2 = new string[6] { "", "", "", "", "", "" };
                        serialPort1.DiscardOutBuffer();
                        serialPort1.DiscardInBuffer();
                        serialPort1.WriteLine("bp");

                        int j = 0;
                        while (true)
                        {
                            try
                            {
                                subs[0] = serialPort1.ReadLine();
                            }
                            catch (TimeoutException) { }
                            tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + subs[0];
                            if (subs[0].Contains("Sent BLE_PAIRING_SUCCESS_SIG from menu"))
                                break;
                            
                            j++;
                            if (j == 6)
                                break;
                        }

                        if (subs[0].Contains("Sent BLE_PAIRING_SUCCESS_SIG from menu"))
                        {
                            tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + "bp disabled";
                            t++;

                            tbpass.BackColor = Color.Green;
                            break;
                        }
                        i++;
                        Thread.Sleep(500);
                        tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + "Attemp" + i.ToString();
                        if (i == 6)
                        {
                            break;
                        }



                    }
                    
                    VV1 = "";
                    i = 0;
                    t= 0;   
                    tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + "Serial Number Eeprom Process" + $"{Environment.NewLine}" + "m017";
                    while (true)
                    {
                        subs = new string[6] { "", "", "", "", "", "" };
                        subs2 = new string[6] { "", "", "", "", "", "" };
                        try
                        {
                            serialPort1.DiscardOutBuffer();
                            serialPort1.DiscardInBuffer();
                            serialPort1.WriteLine("m017");
                            subs[0] = serialPort1.ReadLine();
                            subs[1] = serialPort1.ReadLine();
                            subs[2] = serialPort1.ReadLine();
                        }
                        catch (TimeoutException) { }
                        if (subs[1].Length > 0)
                            subs2 = subs[1].Split(' ');
                        if (subs2.Length > 2)
                        {

                            if (subs2[2].Length >= 9)
                                VV1 = subs2[2].Substring(0, 9);
                            tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + VV1+ $"{Environment.NewLine}"+subs[0] + $"{Environment.NewLine}" + subs[1] + $"{Environment.NewLine}" + subs[2] ;


                        }
                        if(SN == VV1)
                        {
                            i = 7;
                            break;
                        }                 
                        i++;
                        Thread.Sleep(500);
                        tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + "Attemp"+i.ToString();
                        if (i == 6)
                        {
                            break;
                        }
                    }

                    if (SN == VV1 && SN != "")
                    {
                        n = dataGridView1.Rows.Add();
                        dataGridView1.Rows[n].Cells[0].Value = "EEprom";
                        dataGridView1.Rows[n].Cells[1].Value = VV1;
                        dataGridView1.Rows[n].Cells[2].Value = SN;
                        dataGridView1.Rows[n].Cells[3].Value = SN;
                        dataGridView1.Rows[n].Cells[4].Value = "Pass";
                        dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Green;
                        t = t + 1;
                        tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + "Serial Number Eeprom Process" + $"{Environment.NewLine}" + "m017";


                        int FG = 0;
                        //proceso para obtener carga
                        t = 0;
                        i = 0;
                        tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + "Check Fuel Gauge" + $"{Environment.NewLine}" + "m007";
                        while (true)
                        {
                            VV1 = "";
                            subs = new string[6] { "", "", "", "", "", "" };
                            subs2 = new string[6] { "", "", "", "", "", "" };
                            string[] subs3 = new string[6] { "", "", "", "", "", "" };
                            serialPort1.DiscardOutBuffer();
                            serialPort1.DiscardInBuffer();
                            serialPort1.WriteLine("m007");
                            int j = 0;
                            while (j < 5)
                            {

                                try
                                {

                                    subs[0] = serialPort1.ReadLine();

                                }
                                catch (TimeoutException) { }

                                FG = 0;
                                if (subs[0].Length > 0)
                                    subs2 = subs[0].Split(' ');
                                if (subs2.Length < 5 && subs2.Length > 3)
                                {
                                    if (subs2[3].Length >= 5)
                                    {
                                        subs3 = subs2[3].Split('.');
                                        VV1 = subs3[0];
                                        FG = int.Parse(VV1);
                                    }
                                }
                                try
                                {
                                    tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + subs[0] + $"{Environment.NewLine}" + "try1" + j.ToString() + $"{Environment.NewLine}" + VV1 + "vv1";
                                }
                                catch (Exception) { }
                                if (FG > 0 && FG <= 100)
                                {
                                    i = 7;
                                    t = 3;
                                    break;
                                }
                                j = j + 1;
                                //tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + subs[0] + $"{Environment.NewLine}" + "try1" + j.ToString() + $"{Environment.NewLine}" + VV1 + "vv1";


                            }
                            tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + "try2" + i.ToString();
                            if (i == 4)
                                break;
                            if (i == 7)
                                break;
                            i = i + 1;
                        }
                        if (FG >= 95)
                        {
                            t = 3;
                            n = dataGridView1.Rows.Add();
                            dataGridView1.Rows[n].Cells[0].Value = "Fuel Gauge Report";
                            dataGridView1.Rows[n].Cells[1].Value = FG.ToString();
                            dataGridView1.Rows[n].Cells[2].Value = "95";
                            dataGridView1.Rows[n].Cells[3].Value = "100";
                            dataGridView1.Rows[n].Cells[4].Value = "Pass";
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Green;
                            tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + subs[0] + $"{Environment.NewLine}" + subs[1] + $"{Environment.NewLine}" + subs[2] + $"{Environment.NewLine}" + subs[3] + $"{Environment.NewLine}" + FG.ToString();
                        }
                        if (FG < 95)
                        {
                            t = 0;
                            n = dataGridView1.Rows.Add();
                            dataGridView1.Rows[n].Cells[0].Value = "Fuel Gauge Report";
                            dataGridView1.Rows[n].Cells[1].Value = FG.ToString();
                            dataGridView1.Rows[n].Cells[2].Value = "95";
                            dataGridView1.Rows[n].Cells[3].Value = "100";
                            dataGridView1.Rows[n].Cells[4].Value = "Fail";
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;
                        }



                        if(t==3)
                        { 
                        i = 0;
                        VV1 = "";
                        // loop to check Sensor REV ID
                        while (true)
                        {
                            subs = new string[6] { "", "", "", "", "", "" };
                            subs2 = new string[6] { "", "", "", "", "", "" };
                            serialPort1.DiscardOutBuffer();
                            serialPort1.DiscardInBuffer();
                            serialPort1.WriteLine("m084");

                            int j = 0;
                            while (true)
                            {
                                try
                                {
                                    subs[0] = serialPort1.ReadLine();
                                }
                                catch (TimeoutException) { }
                                tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + subs[0];
                                if (subs[0].Contains("Rev G"))
                                    break;
                                if (subs[0].Contains("Rev H"))
                                    break;
                                if (subs[0].Contains("Rev I"))
                                    break;
                                if (subs[0].Contains("Rev J"))
                                    break;
                                if (subs[0].Contains("Rev L"))
                                    break;
                                
                                j++;
                                if (j == 6)
                                    break;
                            }

                            if (subs[0].Contains("Rev L"))
                            {
                                tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + "Nordic Correct";
                                n = dataGridView1.Rows.Add();
                                dataGridView1.Rows[n].Cells[0].Value = "Revision sensor";
                                dataGridView1.Rows[n].Cells[1].Value = "Rev L";
                                dataGridView1.Rows[n].Cells[2].Value = "Rev L";
                                dataGridView1.Rows[n].Cells[3].Value = "Rev L";
                                dataGridView1.Rows[n].Cells[4].Value = "Pass";
                                dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Green;
                                    while (true)
                                    {
                                        subs = new string[6] { "", "", "", "", "", "" };
                                        subs2 = new string[6] { "", "", "", "", "", "" };
                                        serialPort1.DiscardOutBuffer();
                                        serialPort1.DiscardInBuffer();
                                        serialPort1.WriteLine("m058");
                                        j = 0;
                                        while (true)
                                        {
                                            try
                                            {
                                                subs[0] = serialPort1.ReadLine();
                                            }
                                            catch (TimeoutException) { }
                                            tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + subs[0];
                                            if (subs[0].Contains("3.9.0.0"))
                                                break;
                                            j++;
                                            if (j == 6)
                                                break;
                                        }
                                        if (i == 6)
                                        {
                                            t = 89;
                                            break;
                                        }
                                        i++;
                                    }
                                    if (subs[0].Contains("3.9.0.0"))
                                    {
                                        t = 4;
                                        n = dataGridView1.Rows.Add();
                                        dataGridView1.Rows[n].Cells[0].Value = "Shipmode Version";
                                        dataGridView1.Rows[n].Cells[1].Value = "3.9.0.0";
                                        dataGridView1.Rows[n].Cells[2].Value = "3.9.0.0";
                                        dataGridView1.Rows[n].Cells[3].Value = "3.9.0.0";
                                        dataGridView1.Rows[n].Cells[4].Value = "Pass";
                                        dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Green;
                                        tbpass.BackColor = Color.Green;
                                    }
                                    else
                                    {
                                        i = 59;
                                        t = 55;
                                        n = dataGridView1.Rows.Add();
                                        dataGridView1.Rows[n].Cells[0].Value = "Shipmode Version";
                                        dataGridView1.Rows[n].Cells[1].Value = "No Response";
                                        dataGridView1.Rows[n].Cells[2].Value = "3.9.0.0";
                                        dataGridView1.Rows[n].Cells[3].Value = "3.9.0.0";
                                        dataGridView1.Rows[n].Cells[4].Value = "Pass";
                                        dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Green;
                                        tbpass.BackColor = Color.Green;
                                    }
                                        
                                        break;
                            }
                            else if (subs[0].Contains("Rev E"))
                            {
                                t = 99;
                                break;
                            }
                            else if (subs[0].Contains("Rev F"))
                            {
                                t = 99;
                                break;
                            }
                            else if (subs[0].Contains("Rev G"))
                            {
                                t = 99;
                                break;
                            }
                            else if (subs[0].Contains("Rev H"))
                            {
                                t = 99;
                                break;
                            }
                            else if (subs[0].Contains("Rev I"))
                            {
                                t = 99;
                                break;
                            }
                            else if (subs[0].Contains("Rev J"))
                            {
                                t = 99;
                                break;
                            }
                            i++;
                            Thread.Sleep(500);
                            tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + "Attemp" + i.ToString();
                            if (i == 6)
                            {
                                break;
                            }



                        }
                        //after 6 Attemps we assume No Response in Sensor revision
                        if (i == 6)
                        {
                            n = dataGridView1.Rows.Add();
                            dataGridView1.Rows[n].Cells[0].Value = "Sensor revision";
                            dataGridView1.Rows[n].Cells[1].Value = "No response";
                            dataGridView1.Rows[n].Cells[2].Value = "Rev L";
                            dataGridView1.Rows[n].Cells[3].Value = "Rev L";
                            dataGridView1.Rows[n].Cells[4].Value = "Fail";
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;

                            n = dataGridView1.Rows.Add();
                            dataGridView1.Rows[n].Cells[0].Value = "Shipmode Version";
                            dataGridView1.Rows[n].Cells[1].Value = "";
                            dataGridView1.Rows[n].Cells[2].Value = "3.9.0.0";
                            dataGridView1.Rows[n].Cells[3].Value = "3.9.0.0";
                            dataGridView1.Rows[n].Cells[4].Value = "Pass";
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Green;
                        }
                        //this condition if has Rev J 
                        if (t == 99 && i != 6)
                        {
                            tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + "Nordic Correct";
                            n = dataGridView1.Rows.Add();
                            dataGridView1.Rows[n].Cells[0].Value = "Revision Sensor";
                            dataGridView1.Rows[n].Cells[1].Value = "Rev J";
                            dataGridView1.Rows[n].Cells[2].Value = "Rev J";
                            dataGridView1.Rows[n].Cells[3].Value = "Rev J";
                            dataGridView1.Rows[n].Cells[4].Value = "Pass";
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Green;
                            //looop to Check the Shipmode Version if has the REV J
                            while (true)
                            {
                                subs = new string[6] { "", "", "", "", "", "" };
                                subs2 = new string[6] { "", "", "", "", "", "" };
                                serialPort1.DiscardOutBuffer();
                                serialPort1.DiscardInBuffer();
                                serialPort1.WriteLine("m058");
                                    int j = 0;
                                    while (true)
                                    {
                                        try
                                        {
                                            subs[0] = serialPort1.ReadLine();
                                        }
                                        catch (TimeoutException) { }
                                        tbcom.Text = tbcom.Text + $"{Environment.NewLine}" + subs[0];
                                        if (subs[0].Contains("3.8.0.0"))
                                            break;
                                        if (subs[0].Contains("3.7.0.0"))
                                            break;
                                        if (subs[0].Contains("3.5.0.0"))
                                            break;
                                        if (subs[0].Contains("3.6.0.0"))
                                            break;


                                        j++;
                                        if (j == 6)
                                            break;
                                    }
                                    if (subs[0].Contains("3.8.0.0"))
                                    {
                                    n = dataGridView1.Rows.Add();
                                    dataGridView1.Rows[n].Cells[0].Value = "Shipmode Version";
                                    dataGridView1.Rows[n].Cells[1].Value = "3.8.0.0";
                                    dataGridView1.Rows[n].Cells[2].Value = "3.8.0.0";
                                    dataGridView1.Rows[n].Cells[3].Value = "3.8.0.0";
                                    dataGridView1.Rows[n].Cells[4].Value = "Pass";
                                    dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Green;
                                    t = 4;
                                    break;
                                }
                                

                                if (i == 6)
                                {
                                    t = 89;
                                    break;
                                }
                                    i++;
                            }
                            if (t == 89)
                            {
                                MessageBox.Show("Para Actualizar");
                                n = dataGridView1.Rows.Add();
                                dataGridView1.Rows[n].Cells[0].Value = "Shipmode Version";
                                dataGridView1.Rows[n].Cells[1].Value = "3.7.0.0";
                                dataGridView1.Rows[n].Cells[2].Value = "3.8.0.0";
                                dataGridView1.Rows[n].Cells[3].Value = "3.8.0.0";
                                dataGridView1.Rows[n].Cells[4].Value = "Fail";
                                dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;
                                t = 58;
                            }
                        }
                        
                        if (t == 4)
                        {
                            n = dataGridView1.Rows.Add();
                            dataGridView1.Rows[n].Cells[0].Value = "Test Result";
                            dataGridView1.Rows[n].Cells[1].Value = "Pass";
                            dataGridView1.Rows[n].Cells[2].Value = "Pass";
                            dataGridView1.Rows[n].Cells[3].Value = "Pass";
                            dataGridView1.Rows[n].Cells[4].Value = "Pass";
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Green;
                            i = 7;
                            tbpass.Text = "Pass";
                            tbpass.BackColor = Color.Green;
                            panel1.BackColor = Color.Green;
                        }
                        else
                        {

                            n = dataGridView1.Rows.Add();
                            dataGridView1.Rows[n].Cells[0].Value = "Test Result";
                            dataGridView1.Rows[n].Cells[1].Value = "Fail";
                            dataGridView1.Rows[n].Cells[2].Value = "Pass";
                            dataGridView1.Rows[n].Cells[3].Value = "Pass";
                            dataGridView1.Rows[n].Cells[4].Value = "Fail";
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;
                            tbpass.Text = "Fail";
                            tbpass.BackColor = Color.Red;
                            panel1.BackColor = Color.Red;
                        }
                    }
                    else
                    {
                            n = dataGridView1.Rows.Add();
                            dataGridView1.Rows[n].Cells[0].Value = "Revision_ID";
                            dataGridView1.Rows[n].Cells[1].Value = "";
                            dataGridView1.Rows[n].Cells[2].Value = "Pass";
                            dataGridView1.Rows[n].Cells[3].Value = "Pass";
                            dataGridView1.Rows[n].Cells[4].Value = "";
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;
                            n = dataGridView1.Rows.Add();
                            dataGridView1.Rows[n].Cells[0].Value = "Shipmode_Version";
                            dataGridView1.Rows[n].Cells[1].Value = "";
                            dataGridView1.Rows[n].Cells[2].Value = "Pass";
                            dataGridView1.Rows[n].Cells[3].Value = "Pass";
                            dataGridView1.Rows[n].Cells[4].Value = "";
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;
                            n = dataGridView1.Rows.Add();
                            dataGridView1.Rows[n].Cells[0].Value = "Fuel Gauge";
                            dataGridView1.Rows[n].Cells[1].Value = "";
                            dataGridView1.Rows[n].Cells[2].Value = "Pass";
                            dataGridView1.Rows[n].Cells[3].Value = "Pass";
                            dataGridView1.Rows[n].Cells[4].Value = "";
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;

                            n = dataGridView1.Rows.Add();
                            dataGridView1.Rows[n].Cells[0].Value = "Test Result";
                            dataGridView1.Rows[n].Cells[1].Value = "Fuel_Gauge";
                            dataGridView1.Rows[n].Cells[2].Value = "Pass";
                            dataGridView1.Rows[n].Cells[3].Value = "Pass";
                            dataGridView1.Rows[n].Cells[4].Value = "Fail";
                            dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;
                        }
                    }
                    else
                    {

                        tbpass.Text = "Fail";
                        tbpass.BackColor = Color.Red;
                        panel1.BackColor = Color.Red;
                        n = dataGridView1.Rows.Add();
                        dataGridView1.Rows[n].Cells[0].Value = "EEprom";
                        dataGridView1.Rows[n].Cells[1].Value = VV1;
                        dataGridView1.Rows[n].Cells[2].Value = SN;
                        dataGridView1.Rows[n].Cells[3].Value = SN;
                        dataGridView1.Rows[n].Cells[4].Value = "Fail";
                        dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;
                        n = dataGridView1.Rows.Add();
                        dataGridView1.Rows[n].Cells[0].Value = "Fuel Gauge";
                        dataGridView1.Rows[n].Cells[1].Value = "";
                        dataGridView1.Rows[n].Cells[2].Value = "Pass";
                        dataGridView1.Rows[n].Cells[3].Value = "Pass";
                        dataGridView1.Rows[n].Cells[4].Value = "";
                        dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;
                        n = dataGridView1.Rows.Add();
                        dataGridView1.Rows[n].Cells[0].Value = "Revision_ID";
                        dataGridView1.Rows[n].Cells[1].Value = "";
                        dataGridView1.Rows[n].Cells[2].Value = "Pass";
                        dataGridView1.Rows[n].Cells[3].Value = "Pass";
                        dataGridView1.Rows[n].Cells[4].Value = "";
                        dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;
                        n = dataGridView1.Rows.Add();
                        dataGridView1.Rows[n].Cells[0].Value = "Shipmode_Version";
                        dataGridView1.Rows[n].Cells[1].Value = "";
                        dataGridView1.Rows[n].Cells[2].Value = "Pass";
                        dataGridView1.Rows[n].Cells[3].Value = "Pass";
                        dataGridView1.Rows[n].Cells[4].Value = "";
                        dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;
                        n = dataGridView1.Rows.Add();
                        dataGridView1.Rows[n].Cells[0].Value = "Fuel Gauge";
                        dataGridView1.Rows[n].Cells[1].Value = "";
                        dataGridView1.Rows[n].Cells[2].Value = "Pass";
                        dataGridView1.Rows[n].Cells[3].Value = "Pass";
                        dataGridView1.Rows[n].Cells[4].Value = "";
                        dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;
                        
                        n = dataGridView1.Rows.Add();
                        dataGridView1.Rows[n].Cells[0].Value = "Test Result";
                        dataGridView1.Rows[n].Cells[1].Value = "SN Wrong";
                        dataGridView1.Rows[n].Cells[2].Value = "Pass";
                        dataGridView1.Rows[n].Cells[3].Value = "Pass";
                        dataGridView1.Rows[n].Cells[4].Value = "Fail";
                        dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;

                    }
                    
                }
                        
                else
                {

                    tbpass.Text = "Fail";
                    tbpass.BackColor = Color.Red;
                    panel1.BackColor = Color.Red;
                    //adicion de nuevo renglon
                    int n = dataGridView1.Rows.Add();
                    dataGridView1.Rows[n].Cells[0].Value = "NFC Conection";
                    dataGridView1.Rows[n].Cells[1].Value = "Fail";
                    dataGridView1.Rows[n].Cells[2].Value = "Pass";
                    dataGridView1.Rows[n].Cells[3].Value = "Pass";
                    dataGridView1.Rows[n].Cells[4].Value = "Fail";
                    dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;
                    n = dataGridView1.Rows.Add();
                    dataGridView1.Rows[n].Cells[0].Value = "EEprom";
                    dataGridView1.Rows[n].Cells[1].Value = "";
                    dataGridView1.Rows[n].Cells[2].Value = SN;
                    dataGridView1.Rows[n].Cells[3].Value = SN;
                    dataGridView1.Rows[n].Cells[4].Value = "";
                    dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;                    
                    n = dataGridView1.Rows.Add();
                    dataGridView1.Rows[n].Cells[0].Value = "Fuel Gauge";
                    dataGridView1.Rows[n].Cells[1].Value = "";
                    dataGridView1.Rows[n].Cells[2].Value = "Pass";
                    dataGridView1.Rows[n].Cells[3].Value = "Pass";
                    dataGridView1.Rows[n].Cells[4].Value = "";
                    dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;
                    n = dataGridView1.Rows.Add();
                    dataGridView1.Rows[n].Cells[0].Value = "Revision_ID";
                    dataGridView1.Rows[n].Cells[1].Value = "";
                    dataGridView1.Rows[n].Cells[2].Value = "Pass";
                    dataGridView1.Rows[n].Cells[3].Value = "Pass";
                    dataGridView1.Rows[n].Cells[4].Value = "";
                    dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;
                    n = dataGridView1.Rows.Add();
                    dataGridView1.Rows[n].Cells[0].Value = "Shipmode_Version";
                    dataGridView1.Rows[n].Cells[1].Value = "";
                    dataGridView1.Rows[n].Cells[2].Value = "Pass";
                    dataGridView1.Rows[n].Cells[3].Value = "Pass";
                    dataGridView1.Rows[n].Cells[4].Value = "";
                    dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;
                    n = dataGridView1.Rows.Add();
                    dataGridView1.Rows[n].Cells[0].Value = "Test Result";
                    dataGridView1.Rows[n].Cells[1].Value = "NFC_BAD_Comunication";
                    dataGridView1.Rows[n].Cells[2].Value = "Pass";
                    dataGridView1.Rows[n].Cells[3].Value = "Pass";
                    dataGridView1.Rows[n].Cells[4].Value = "Fail";
                    dataGridView1.Rows[n].DefaultCellStyle.BackColor = Color.Red;
                }
                serialPort1.Close();
                tbsn.Text = "";
                tbsn.Select();
            }
            else
            {
                MessageBox.Show("Please scan the serial number", "Salir", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
            }
            tbsn.Text = "";
            string pathString = @"c:\Report_VIP\VIP_Shipmode_" + DateTime.Now.ToString("MM_dd_yyyy_HH_mm") +"_" +SN+".csv";
            string pathString2 = @"c:\Report_VIP\log_VIP_Shipmode_" + DateTime.Now.ToString("MM_dd_yyyy_HH_mm") +"_" +SN + ".txt";
            string log=tbcom.Text;
            string csvt = "";
            string csv = "";
            sw1.Stop(); // Detener la medición.
            tbtime.Text = sw1.Elapsed.ToString("hh\\:mm\\:ss");
            if (dataGridView1.Rows.Count > 1)
            {
            
                
                for (i = 0; i < (dataGridView1.Rows.Count-1); i++)
                {
                    csvt = csvt + dataGridView1.Rows[i].Cells[0].Value + ",";
                    csv = csv + dataGridView1.Rows[i].Cells[1].Value + ",";                   
                }
            }
            csvt = "TEST_PROGRAM_VERSION" + "," + "TEST_START_TIME" + "," + "TEST_END_TIME" + "," + "TEST_DURATION" + "," + "SERIAL_NUMBER" + "," + "TEST_RESULT" + "," + csvt + "WORKSTATION" + "," + "EMPLOYEE_ID" + "\n";
            csv = csvt + "QMSM_Test_Shipmode_03_14_2023" + "," + Teststart + "," + DateTime.Now.ToString("MM/dd/yyyy HH\\:mm\\:ss") + "," + sw1.Elapsed.ToString("hh\\:mm\\:ss") + "" + "," + SN + "," + dataGridView1.Rows[3].Cells[1].Value + "," + csv + "Test Station #1" + "," + "140183A" + "\n";





            using (StreamWriter sw = File.CreateText(pathString))
            {
                sw.Write(csv);
            }
            File.WriteAllText(pathString2, log);
            
 
            
        }
    

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("¿Desea Salir?", "Salir", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.Yes)
                Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            string dia = Application.CurrentCulture.DateTimeFormat.GetDayName(DateTime.Now.DayOfWeek).ToString();
            diadelasemana = dia;
            string day = DateTime.Now.Day.ToString("00");
            dianum = day;
            string mes1 = DateTime.Now.Month.ToString("00");
            mes = mes1;
            string hora1 = DateTime.Now.Hour.ToString("00");
            hora = hora1;

            string minutos1 = DateTime.Now.Minute.ToString("00");
            minutos = minutos1;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Aspose.Cells.Workbook wb1 = new Aspose.Cells.Workbook();
        }
    }
}
