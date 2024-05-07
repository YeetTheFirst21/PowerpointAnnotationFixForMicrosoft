using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;

namespace Wpf_annotate
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.KeyDown += new KeyEventHandler(MainWindow_KeyDown);
        }

        [DllImport("user32.dll")]
        public static extern int SetForegroundWindow(IntPtr hWnd);

        [STAThread]
        private void MainWindow_KeyDown(object sender, KeyEventArgs e)
        {
            //if c is pressed, clean the inkcanvas
            if (e.Key == Key.C)
            {
                //clear all text written in inkcanvas
                inkCanvas1.Strokes.Clear();
            }else if(e.Key == Key.Q)
            {
                //close the app
                this.Close();
            }
            
            //else forward all input to the other app running behin it:
            else
            {
                if (e.Key == Key.Right)
                {
                    Process[] processes = Process.GetProcesses();

                    //get only processes that has window title PowerPoint
                    processes = processes.Where(p => p.MainWindowTitle.Contains("PowerPoint")).ToArray();

                    foreach (Process proc in processes)
                    {
                        SetForegroundWindow(proc.MainWindowHandle);
                        //send right arrow
                        System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                        //sleep for 10 ms
                        System.Threading.Thread.Sleep(10);
                        //send alt+tab
                        //System.Windows.Forms.SendKeys.SendWait("%{TAB}");
                        //make this app the active window
                        SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);

                        
                    }

                }else if(e.Key== Key.Left)
                {
                    Process[] processes = Process.GetProcesses();
                    //get only processes that has window title PowerPoint
                    processes = processes.Where(p => p.MainWindowTitle.Contains("PowerPoint")).ToArray();
                    foreach (Process proc in processes)
                    {
                        SetForegroundWindow(proc.MainWindowHandle);
                        //send left arrow
                        System.Windows.Forms.SendKeys.SendWait("{LEFT}");
                        //sleep for 10 ms
                        System.Threading.Thread.Sleep(10);
                        //send alt+tab
                        //System.Windows.Forms.SendKeys.SendWait("%{TAB}");
                        //make this app the active window
                        SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);
                    }
                }
                e.Handled= false;
            }
        }

        private void Window_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            Process[] processes = Process.GetProcesses();
            //get only processes that has window title PowerPoint
            processes = processes.Where(p => p.MainWindowTitle.Contains("PowerPoint")).ToArray();
            foreach (Process proc in processes)
            {
                SetForegroundWindow(proc.MainWindowHandle);
                //send left arrow
                System.Windows.Forms.SendKeys.SendWait("{LEFT}");
                //sleep for 10 ms
                System.Threading.Thread.Sleep(10);
                //send alt+tab
                //System.Windows.Forms.SendKeys.SendWait("%{TAB}");
                //make this app the active window
                SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);
            }

            //if the mouse wheel is scrolled down, send right arrow key
            if (e.Delta < 0)
            {
                foreach (Process proc in processes)
                {
                    SetForegroundWindow(proc.MainWindowHandle);
                    //send right arrow
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    //sleep for 10 ms
                    System.Threading.Thread.Sleep(10);
                    //send alt+tab
                    //System.Windows.Forms.SendKeys.SendWait("%{TAB}");
                    //make this app the active window
                    SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);
                }
            }
            //if the mouse wheel is scrolled up, send left arrow key
            else
            {
                foreach (Process proc in processes)
                {
                    SetForegroundWindow(proc.MainWindowHandle);
                    //send left arrow
                    System.Windows.Forms.SendKeys.SendWait("{LEFT}");
                    //sleep for 10 ms
                    System.Threading.Thread.Sleep(10);
                    //send alt+tab
                    //System.Windows.Forms.SendKeys.SendWait("%{TAB}");
                    //make this app the active window
                    SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);
                }
            }
        }
    }
}
