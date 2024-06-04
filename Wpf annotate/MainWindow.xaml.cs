using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;
using Button = System.Windows.Controls.Button;
using HorizontalAlignment = System.Windows.HorizontalAlignment;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using KeyEventHandler = System.Windows.Input.KeyEventHandler;
using Label = System.Windows.Controls.Label;
using MessageBox = System.Windows.MessageBox;
using TextBox = System.Windows.Controls.TextBox;

namespace Wpf_annotate
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private System.Windows.Forms.ColorDialog colorDialog = new System.Windows.Forms.ColorDialog();
        private readonly Microsoft.Office.Interop.PowerPoint.Application PPT = 
            new Microsoft.Office.Interop.PowerPoint.Application();

        private Microsoft.Office.Interop.PowerPoint.SlideShowWindow getSlideShowWindow()
        {
            foreach (SlideShowWindow window in PPT.SlideShowWindows)
            {
                return window;
            }
            return null;
        }

        public MainWindow()
        {
            InitializeComponent();
            this.KeyDown += new KeyEventHandler(MainWindow_KeyDown);
        }

        [DllImport("user32.dll")]
        public static extern int SetForegroundWindow(IntPtr hWnd);
        // https://stackoverflow.com/a/3744720/8302811
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(HandleRef hWnd, [In, Out] ref RECT rect);
        // https://stackoverflow.com/q/39458046/8302811
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        private const int SW_MAXIMIZE = 3;
        private const int SW_MINIMIZE = 6;

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
                if (PPT.SlideShowWindows.Count == 0)
                {
                    e.Handled = false;
                    return;
                }

                if (e.Key == Key.Left)
                {
                    foreach (Process proc in PPT.SlideShowWindows)
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
                    foreach (Process proc in PPT.SlideShowWindows)
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

            }else if (e.Key == Key.P)
            {
                //open color picker
                colorDialog.ShowDialog();
                //set the color of the inkcanvas to the color selected
                System.Drawing.Color color = colorDialog.Color;
                inkCanvas1.DefaultDrawingAttributes.Color = System.Windows.Media.Color.FromArgb(color.A, color.R, color.G, color.B);
            }//if o is pressed, select pen width:
            else if (e.Key == Key.O)
            {
                //open a textbox and a button in a border to enter the pen width with the current pen width as default value:
                Border border = new Border();
                border.Width = 200;
                border.Height = 100;
                border.BorderBrush = System.Windows.Media.Brushes.Black;
                border.BorderThickness = new Thickness(5);
                border.CornerRadius = new CornerRadius(20);
                border.HorizontalAlignment = HorizontalAlignment.Center;
                border.VerticalAlignment = VerticalAlignment.Center;
                border.Background = System.Windows.Media.Brushes.White;



                TextBox textBox = new TextBox();
                textBox.HorizontalAlignment = HorizontalAlignment.Center;
                textBox.VerticalAlignment = VerticalAlignment.Center;
                textBox.Width = 50;
                textBox.Height = 20;
                textBox.Margin = new Thickness(0, 0, 0, 20);

                //when enter is pressed, set the pen width to the value in the textbox
                textBox.KeyDown += (object senderrr, KeyEventArgs eee) =>
                {
                    if (eee.Key == Key.Enter)
                    {
                        savePencil(border, textBox);
                    }
                };

                //add a label to grid2 in its center to show the user what to do
                Label label = new Label();
                label.Content = "Enter Pen Width:";
                label.HorizontalAlignment = HorizontalAlignment.Center;
                label.VerticalAlignment = VerticalAlignment.Top;
                label.Margin = new Thickness(0, 0, 0, 0);
                label.Width = 150;
                label.Height = 27;
                label.FontSize = 12;
                label.FontWeight = FontWeights.Bold;
                label.Foreground = System.Windows.Media.Brushes.Black;
                label.Background = System.Windows.Media.Brushes.White;

                Button button = new Button();
                button.Content = "Set";
                button.HorizontalAlignment = HorizontalAlignment.Center;
                button.VerticalAlignment = VerticalAlignment.Bottom;
                button.Width = 50;
                button.Height = 20;
                button.Margin = new Thickness(0, 20, 0, 0);
                button.Click += (object senderr, RoutedEventArgs ee) =>
                {
                    savePencil(border, textBox);
                };
                
                //border only accepts one child, so we add a grid to it
                Grid grid2 = new Grid();
                grid2.Children.Add(textBox);
                grid2.Children.Add(button);
                grid2.Children.Add(label);
                border.Child = grid2;
                border.Visibility = Visibility.Visible;
                grid1.Children.Add(border);

                textBox.Focus();
                e.Handled = true;
                textBox.Text = inkCanvas1.DefaultDrawingAttributes.Width.ToString();
                //make the input index at the end:
                textBox.SelectionStart = textBox.Text.Length;

            }
            else if (e.Key == Key.H)
            {
                // Hide window
                WindowState = WindowState.Minimized;
                if (PPT.SlideShowWindows.Count > 0)
                {
                    ShowWindow((IntPtr) getSlideShowWindow().HWND, SW_MINIMIZE);
                }
            }
            else
            {
                e.Handled = false;
            }
        }
        private void savePencil(Border border, TextBox textBox)
        {
            double val;
            if (double.TryParse(textBox.Text, out val))
            {
                inkCanvas1.DefaultDrawingAttributes.Width = val;
                inkCanvas1.DefaultDrawingAttributes.Height = val;
                border.Visibility = Visibility.Hidden;
            }
            else
            {
                MessageBox.Show("Please enter a valid number", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Window_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (PPT.SlideShowWindows.Count == 0)
            {
                e.Handled = false;
                return;
            }
            //if the mouse wheel is scrolled down, send right arrow key
            if (e.Delta < 0)
            {
                foreach (Process proc in PPT.SlideShowWindows)
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
                foreach (Process proc in PPT.SlideShowWindows)
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

        // https://stackoverflow.com/a/3744720/8302811
        public static bool IsFullScreen(Process process, Screen screen = null)
        {
            if (screen == null)
            {
                screen = Screen.PrimaryScreen;
            }
            RECT rect = new RECT();
            GetWindowRect(new HandleRef(null, process.MainWindowHandle), ref rect);
            return rect.right - rect.left >= screen.Bounds.Width && rect.bottom - rect.top >= screen.Bounds.Height;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            BringPowerpointForward();
        }

        private void BringPowerpointForward()
        {
            if (PPT.SlideShowWindows.Count > 0)
            {
                SetForegroundWindow((IntPtr) getSlideShowWindow().HWND);
            }
        }

        private void Window_StateChanged(object sender, EventArgs e)
        {
            if (WindowState == WindowState.Maximized)
            {
                BringPowerpointForward();
            }
        }
    }
}
