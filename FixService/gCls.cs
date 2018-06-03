using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Net.NetworkInformation;

namespace IdFix
{
    public static class gCls
    {
        



        public static void update_path_var()
        {
            try
            {
                string gv = Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Machine);
                string xPath = Application.StartupPath + "\\idex";
                if (gv != "")
                {
                    string[] allv = gv.Split(';');
                    bool pathFound = false;
                    foreach (string s in allv)
                    {
                        if (s == xPath)
                        {
                            pathFound = true;
                        }
                    }
                    if (pathFound == false)
                    {
                        Environment.SetEnvironmentVariable("PATH", gv + ";" + xPath, EnvironmentVariableTarget.Machine);
                    }

                }


            }
            catch (Exception erd)
            {
                show_error(erd.Message.ToString() + "\n\nRun as administrator to registry access");
                return;
            }

        }

        public static string HexConverter(System.Drawing.Color c)
        {
            String rtn = String.Empty;
            try
            {
                rtn = "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
            }
            catch (Exception ex)
            {
                //doing nothing
            }

            return rtn;
        }

        public static string Text2HexaConversion(string text)
        {
            char[] chars = text.ToCharArray();
            StringBuilder result = new StringBuilder(text.Length + (int)(text.Length * 0.1));

            foreach (char c in chars)
            {
                int value = (int)(c);
                string fhvalue = value.ToString("X");
                int dvalue = int.Parse(fhvalue, System.Globalization.NumberStyles.HexNumber);
                string hvalue = dvalue.ToString();

                //if (hvalue.Length == 2) {
                //    hvalue = "00" + hvalue;
                //}
                //else if (hvalue.Length == 3) {
                //    hvalue = "0" + hvalue;
                //}

                if (value > 127)
                    result.AppendFormat("&#{0};", hvalue);
                else
                    result.Append(c);
            }

            return result.ToString();
        }

    

        public static void show_error(string msg)
        {
            try
            {

                ComponentFactory.Krypton.Toolkit.KryptonMessageBox.Show(msg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            catch { }

        }
        public static void show_message(string msg)
        {
            try
            {
                ComponentFactory.Krypton.Toolkit.KryptonMessageBox.Show(msg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch { }

        }

      
       

      

        public static string GetMacAddress()
        {
            string macAddresses = "";
            foreach (NetworkInterface nic in NetworkInterface.GetAllNetworkInterfaces())
            {
                // Only consider Ethernet network interfaces, thereby ignoring any
                // loopback devices etc.
                if (nic.NetworkInterfaceType != NetworkInterfaceType.Ethernet) continue;
                if (nic.OperationalStatus == OperationalStatus.Up)
                {
                    macAddresses += nic.GetPhysicalAddress().ToString();
                    break;
                }
            }
            return macAddresses;
        }

       
       

     

        public static bool FindAndKillProcess(string name)
        {
            try
            {
                bool killp = false;

                foreach (Process clsProcess in Process.GetProcesses())
                {

                    if (clsProcess.ProcessName.StartsWith(name))
                    {
                        clsProcess.Kill();

                        //process killed, return true
                        killp = true;
                    }
                }

                return killp;
            }
            catch { return false; }
        }

        public static void get_image_size(string imgPath, ref int iWidth, ref int iHeight)
        {

            try
            {
                System.Drawing.Image imgBox = System.Drawing.Image.FromFile(imgPath);
                iWidth = imgBox.Size.Width;
                iHeight = imgBox.Size.Height;
                imgBox.Dispose();
            }
            catch { }

        }

       public static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);
            DirectoryInfo[] dirs = dir.GetDirectories();

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }

            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDirName, file.Name);
                file.CopyTo(temppath, false);
            }

            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }

    }

    public class Updf
    {

        public string b_rootpath { get; set; }
        public string b_outpath { get; set; }

        public string b_filenamewoe { get; set; }
        public string b_title { get; set; }
        public string b_author { get; set; }
        public string b_pdffile { get; set; }
        public string b_coverpath { get; set; }

        public string b_dpi { get; set; }



        public Updf(string i_filenamewoe, string i_title, string i_author, string i_pdffile, string i_rootpath, string i_outpath, string i_coverpath,string i_dpi)
        {
            try
            {


                b_filenamewoe = i_filenamewoe;
                b_title = i_title;
                b_author = i_author;
                b_pdffile = i_pdffile;
                b_rootpath = i_rootpath;
                b_outpath = i_outpath;
                b_coverpath = i_coverpath;
                b_dpi = i_dpi;
            }
            catch { }
        }
    }
}
