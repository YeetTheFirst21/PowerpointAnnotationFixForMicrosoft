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
        Process[] processes = new Process[0];
        public MainWindow()
        {
            InitializeComponent();
            this.KeyDown += new KeyEventHandler(MainWindow_KeyDown);
            findPowerPoint();
        }


        private void findPowerPoint()
        {
            //get all processes
            processes = Process.GetProcesses();
            /*
            foreach (Process process in processes)
            {
                if (!String.IsNullOrEmpty(process.MainWindowTitle))
                {
                    Console.WriteLine("Process: {0} ID: {1} Window title: {2}", process.ProcessName, process.Id, process.MainWindowTitle);
                }
            }
            */

            //get only processes that has window title PowerPoint
            processes = processes.Where(p => p.MainWindowTitle.Contains("PowerPoint Slide Show")).ToArray();
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
            }else if(e.Key == Key.Right || e.Key == Key.Left)
            {
                

                //need to recall this as if you alttab into Powerpoints main app, the process will be lost :(
                //findPowerPoint();
                //no more, found a workaround with the or: :)
                if(processes.Length < 1 || processes[0].HasExited)
                {
                    findPowerPoint();
                    if(processes.Length < 1)
                    {
                        MessageBox.Show("Powerpoint not Found!\nYou might try restaring the presentation and not pressing main powerpoint tab back.","Error",MessageBoxButton.OK,MessageBoxImage.Error);
                        e.Handled = false;
                        return;
                    }
                }
                

                if(e.Key == Key.Left)
                {
                    foreach (Process proc in processes)
                    {
                        SetForegroundWindow(proc.MainWindowHandle);
                        //send right arrow after sleeping for 2ms
                        System.Threading.Thread.Sleep(2);
                        System.Windows.Forms.SendKeys.SendWait("{LEFT}");
                        //sleep for 2 ms
                        System.Threading.Thread.Sleep(2);
                        //make this app the active window
                        SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);
                    }
                }
                else
                {
                    foreach (Process proc in processes)
                    {
                        SetForegroundWindow(proc.MainWindowHandle);
                        //send right arrow after sleeping for 2ms
                        System.Threading.Thread.Sleep(2);
                        System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                        //sleep for 2 ms
                        System.Threading.Thread.Sleep(2);
                        //make this app the active window
                        SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);
                    }
                }

            }
            else
            {
                e.Handled= false;
            }
        }

        private void Window_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (processes.Length < 1 || processes[0].HasExited)
            {
                findPowerPoint();
                if (processes.Length < 1)
                {
                    MessageBox.Show("Powerpoint not Found!\nYou might try restaring the presentation and not pressing main powerpoint tab back.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    e.Handled = false;
                    return;
                }
            }
            //if the mouse wheel is scrolled down, send right arrow key
            if (e.Delta < 0)
            {
                foreach (Process proc in processes)
                {
                    SetForegroundWindow(proc.MainWindowHandle);
                    //send right arrow after sleeping for 2ms
                    System.Threading.Thread.Sleep(2);
                    System.Windows.Forms.SendKeys.SendWait("{RIGHT}");
                    //sleep for 2 ms
                    System.Threading.Thread.Sleep(2);
                    //make this app the active window
                    SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);
                }
            }
            //if the mouse wheel is scrolled up, send left arrow key
            else if (e.Delta > 0)
            {
                foreach (Process proc in processes)
                {
                    SetForegroundWindow(proc.MainWindowHandle);
                    //send left arrow after sleeping for 2ms
                    System.Threading.Thread.Sleep(2);
                    System.Windows.Forms.SendKeys.SendWait("{LEFT}");
                    //sleep for 2 ms
                    System.Threading.Thread.Sleep(2);
                    //make this app the active window
                    SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle);
                }
            }
        }
    }
}
