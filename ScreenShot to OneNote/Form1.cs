using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ScreenShot_to_OneNote
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            getProcess();
            if (screenLength == 1){
                //screen1
                this.Size = new Size(260, 310);
            }
            else{
                //screen2
                comboBox1.Visible = false;
            }
            
        }
        //set specific foreground app as active
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        String toReturn = null;
         int screenLength = Screen.AllScreens.Length;

        private void button1_Click(object sender, EventArgs e){                    
            if (screenLength == 1){
                //screen 1 only block          
                if (comboBox1.Text == ""){
                    MessageBox.Show("Select where the active foreground will return");
                    return;
                }         
                SendKeys.SendWait("{PRTSC}");            
                pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                pictureBox1.Image = Clipboard.GetImage();
                pasteToOneNoteSingleScreen();            
                return;
            }

            //else multiplse screen block
            int screenNumber = 1;
            int targetScreen = 2;
            foreach (Screen screen in Screen.AllScreens){
                if (screenNumber == targetScreen){
                Bitmap screenshot = new Bitmap(screen.Bounds.Width, screen.Bounds.Height, PixelFormat.Format32bppArgb);
                Graphics memoryGraphics = Graphics.FromImage(screenshot);
                memoryGraphics.CopyFromScreen(screen.Bounds.X, screen.Bounds.Y, 0, 0, screen.Bounds.Size, CopyPixelOperation.SourceCopy);                       
                pictureBox1.Image = screenshot;
                Clipboard.SetImage(screenshot);
                pasteToOneNoteMultiScreen();     
                }
                screenNumber++;
            }
        }

        public void pasteToOneNoteMultiScreen(){
                var prc = Process.GetProcessesByName("onenote");
                if (prc.Length > 0){
                    SetForegroundWindow(prc[0].MainWindowHandle);
                    Thread.Sleep(320);
                    SendKeys.SendWait("^{v}");
                    prc = Process.GetProcessesByName("Screenshot to OneNote");
                    if (prc.Length > 0)                 
                        SetForegroundWindow(prc[0].MainWindowHandle);                                                  
                }
            
        }

        public void pasteToOneNoteSingleScreen() {
                var prc = Process.GetProcessesByName("onenote");
                if (prc.Length > 0){
                    SetForegroundWindow(prc[0].MainWindowHandle);
                    Thread.Sleep(700);
                    SendKeys.SendWait("^{v}");
                    Thread.Sleep(700);
                    prc = Process.GetProcessesByName(toReturn);
                    if (toReturn.Equals("chrome") || toReturn.Equals("msedge")){
                        int i = 0;
                        Process[] processes = Process.GetProcessesByName(toReturn);                  
                        foreach (var process in Process.GetProcessesByName(toReturn)){
                            SetForegroundWindow(processes[i].MainWindowHandle);
                            i++;
                        }          
                    }
                    else if (prc.Length > 0)
                        SetForegroundWindow(prc[0].MainWindowHandle);
                    prc = Process.GetProcessesByName("Screenshot to OneNote");
                    if (prc.Length > 0)
                        SetForegroundWindow(prc[0].MainWindowHandle);
            }
             
        }

        public void getProcess() {
            List<string> arrList = new List<string>();
            Process[] processCollection = Process.GetProcesses();
            foreach (Process p in processCollection){              
                if(!arrList.Contains(p.ProcessName))
                arrList.Add(p.ProcessName);
            }
           
            foreach(string s in arrList){
                comboBox1.Items.Add(s);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e){
            toReturn = comboBox1.Text;
        }

        private void comboBox1_TextUpdate(object sender, EventArgs e){
            toReturn = comboBox1.Text;
        }

    }
}
