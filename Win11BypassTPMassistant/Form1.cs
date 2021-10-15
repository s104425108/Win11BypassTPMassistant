using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management;
using System.Management.Automation.Runspaces;
using System.IO;
using System.Threading;
using System.Management.Automation;
using System.Collections.ObjectModel;
using System.Diagnostics;

namespace Win11BypassTPMassistant
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void guna2PictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Form.CheckForIllegalCrossThreadCalls = false;
            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync();
            }
        }
        static ulong GetTotalMemoryInBytes()
        {
            return new Microsoft.VisualBasic.Devices.ComputerInfo().TotalPhysicalMemory;
        }


        private static void GetComponent(Guna.UI2.WinForms.Guna2TextBox textBox, string hwclass, string syntax)
        {
            ManagementObjectSearcher mos = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM " + hwclass);
            foreach(ManagementObject mj in mos.Get())
            {
                textBox.Text = $"{Convert.ToString(mj[syntax])}";
                //Console.WriteLine(Convert.ToString(mj[syntax]));
            }
        }

        private static void GetComponent(Label text, string hwclass, string syntax)
        {
            ManagementObjectSearcher mos = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM " + hwclass);
            foreach (ManagementObject mj in mos.Get())
            {
                text.Text = $"您的目前狀態 : {Convert.ToString(mj[syntax])}";
                //Console.WriteLine(Convert.ToString(mj[syntax]));
            }
        }

        async void GreatBigMethod()
        {
            while (!File.Exists($"{Application.StartupPath}\\dxdiag.txt"))
            {
                await Task.Delay(20);
            }

            string[] strs = File.ReadAllLines($"{Application.StartupPath}\\dxdiag.txt");
            var s = strs[16].Split(':');
            DirectX.Text = s[1];

            if (DirectX.Text.Contains("12"))
            {
                Jball.FillColor = Color.Lime;
                Jball.FillColor2 = Color.Green;
            }
            else
            {
                Jball.FillColor = Color.Red;
                Jball.FillColor2 = Color.Maroon;
            }

            if (Aball.FillColor == Color.Lime &&
    Bball.FillColor == Color.Lime &&
    Cball.FillColor == Color.Lime &&
    Dball.FillColor == Color.Lime &&
    Eball.FillColor == Color.Lime &&
    Fball.FillColor == Color.Lime &&
    Gball.FillColor == Color.Lime &&
    Hball.FillColor == Color.Lime &&
    Iball.FillColor == Color.Lime)
            {
                Totalball.FillColor = Color.Lime;
                Totalball.FillColor2 = Color.Green;
                label6.Text = "檢查狀態 : 準備就緒";
            }
            else
            {
                Totalball.FillColor = Color.Red;
                Totalball.FillColor2 = Color.Maroon;
                label6.Text = "檢查狀態 : 有問題";
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            Task task = Task.Run(() => {
                GreatBigMethod();
            });

            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();
            Pipeline pipeline = runspace.CreatePipeline();
            pipeline.Commands.AddScript("Get-Tpm");
            pipeline.Commands.Add("Out-String");
            Collection<PSObject> result = pipeline.Invoke();
            
            StringBuilder stringBuilder = new StringBuilder();
            foreach(PSObject pSObject in result)
            {
                stringBuilder.AppendLine(pSObject.ToString());
            }
            var d = stringBuilder.ToString().Split(':');
            try
            {
                tpmSup.Text += d[2].Substring(0, 6);

                if (tpmSup.Text.Contains("False"))
                {
                    Iball.FillColor = Color.Red;
                    Iball.FillColor2 = Color.Maroon;
                }
                if (tpmSup.Text.Contains("True"))
                {
                    Iball.FillColor = Color.Lime;
                    Iball.FillColor2 = Color.Green;
                }
            }
            catch 
            { 
                tpmSup.Text = stringBuilder.ToString();

                Iball.FillColor = Color.Red;
                Iball.FillColor2 = Color.Maroon;
            }


            backgroundWorker1.ReportProgress(5);
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            stringBuilder.Clear();
            pipeline = runspace.CreatePipeline();

            pipeline.Commands.AddScript("$env:firmware_type");
            
            result = pipeline.Invoke();

            foreach (PSObject pSObject in result)
            {
                stringBuilder.AppendLine(pSObject.ToString());
            }
            bootMode.Text = stringBuilder.ToString();

            if (!bootMode.Text.Contains("UEFI"))
            {
                Aball.FillColor = Color.Red;
                Aball.FillColor2 = Color.Maroon;
            }
            else
            {
                Aball.FillColor = Color.Lime;
                Aball.FillColor2 = Color.Green;
            }

            backgroundWorker1.ReportProgress(10);

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            stringBuilder.Clear();
            pipeline = runspace.CreatePipeline();

            pipeline.Commands.AddScript($"Dxdiag /t {Application.StartupPath}\\dxdiag.txt");

            pipeline.Commands.AddScript("Confirm-SecureBootUEFI");

            try
            {
                result = pipeline.Invoke();

                runspace.Close();

                foreach (PSObject pSObject in result)
                {
                    stringBuilder.AppendLine(pSObject.ToString());
                }
                secureBoot.Text = stringBuilder.ToString();

                if (secureBoot.Text.Contains("False"))
                {
                    Gball.FillColor = Color.Red;
                    Gball.FillColor2 = Color.Maroon;
                }
                if (secureBoot.Text.Contains("True"))
                {
                    Gball.FillColor = Color.Lime;
                    Gball.FillColor2 = Color.Green;
                }
            }
            catch (Exception ex) 
            { 
                if (tpmSup.Text.Contains("執行該命令需要管理員權限")) 
                { 
                    secureBoot.Text = tpmSup.Text; 
                }
                else 
                { 
                    secureBoot.Text = ex.ToString(); 
                }
                Gball.FillColor = Color.Red;
                Gball.FillColor2 = Color.Maroon;
            }

            GetComponent(label3, "Win32_OperatingSystem", "Caption");
            OSball.FillColor = Color.Lime;
            OSball.FillColor2 = Color.Green;
            backgroundWorker1.ReportProgress(15);

            GetComponent(OSArchitecture, "Win32_OperatingSystem", "OSArchitecture");
            if (OSArchitecture.Text.Contains("64"))
            {
                Kball.FillColor = Color.Lime;
                Kball.FillColor2 = Color.Green;
            }
            else
            {
                Kball.FillColor = Color.Red;
                Kball.FillColor2 = Color.Maroon;
            }

            backgroundWorker1.ReportProgress(20);

            GetComponent(coreCount, "Win32_Processor", "NumberOfCores");
            if(int.TryParse(coreCount.Text, out int a))
            {
                if(a < 2)
                {
                    Cball.FillColor = Color.Red;
                    Cball.FillColor2 = Color.Maroon;
                }
                else
                {
                    Cball.FillColor = Color.Lime;
                    Cball.FillColor2 = Color.Green;
                }
            }else
            {
                Cball.FillColor = Color.Yellow;
                Cball.FillColor2 = Color.Goldenrod;
            }
            backgroundWorker1.ReportProgress(30);

            GetComponent(cpuName, "Win32_Processor", "Name");
            Bball.FillColor = Color.Lime;
            Bball.FillColor2 = Color.Green;
            backgroundWorker1.ReportProgress(40);

            GetComponent(cpuFrequence, "Win32_Processor", "MaxClockSpeed");
            if (int.TryParse(cpuFrequence.Text, out int b))
            {
                if(b < 1000)
                {
                    Dball.FillColor = Color.Red;
                    Dball.FillColor2 = Color.Maroon;
                }
                else
                {
                    Dball.FillColor = Color.Lime;
                    Dball.FillColor2 = Color.Green;
                }
            }
            else
            {
                Dball.FillColor = Color.Yellow;
                Dball.FillColor2 = Color.Goldenrod;
            }
            cpuFrequence.Text += " MHz";
            backgroundWorker1.ReportProgress(50);

            GetComponent(diskSpace, "Win32_DiskDrive", "Size");
            backgroundWorker1.ReportProgress(60);

            DriveInfo[] allDrives = DriveInfo.GetDrives();
            diskSpace.Text = $"{allDrives[0].AvailableFreeSpace / 1024 / 1024} MB";
            if(allDrives[0].AvailableFreeSpace / 1024 / 1024 < 65536)
            {
                Hball.FillColor = Color.Red;
                Hball.FillColor2 = Color.Maroon;
            }
            else
            {
                Hball.FillColor = Color.Lime;
                Hball.FillColor2 = Color.Green;
            }

            diskParticial.Text = $"{allDrives[0].DriveFormat}";
            Eball.FillColor = Color.Lime;
            Eball.FillColor2 = Color.Green;


            backgroundWorker1.ReportProgress(80);
            ram.Text = $"{GetTotalMemoryInBytes() / 1024 / 1024 / 1024 + 1} GB";
            if (GetTotalMemoryInBytes() / 1024 / 1024 / 1024 + 1 <= 4)
            {
                Hball.FillColor = Color.Red;
                Hball.FillColor2 = Color.Maroon;
            }
            else
            {
                Fball.FillColor = Color.Lime;
                Fball.FillColor2 = Color.Green;
            }

            backgroundWorker1.ReportProgress(100);

        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            guna2ProgressBar1.Value = e.ProgressPercentage;
            label21.Text = $"{e.ProgressPercentage} %";
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            Process.Start("https://mega.nz/file/8csQjL6B#dgNhS74op4DzQYT9TvkYfBI_Z5X2JEWpIkf5BvbOq4E");
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            Process.Start("https://mega.nz/file/8csQjL6B#dgNhS74op4DzQYT9TvkYfBI_Z5X2JEWpIkf5BvbOq4E");
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            Process.Start("https://mega.nz/file/8csQjL6B#dgNhS74op4DzQYT9TvkYfBI_Z5X2JEWpIkf5BvbOq4E");
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {

        }

        private void guna2Button6_Click(object sender, EventArgs e)
        {

        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            Process.Start("https://www.paypal.com/paypalme/81dooms42/100");
        }

    }
}
