using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Internal;
using System.Threading;
using System.Collections.ObjectModel;
using System.Diagnostics;
using DangDangAutoPal.Models;
using Microsoft.Win32;
using DangDangAutoPal.Properties;
using logger;


namespace DangDangAutoPal
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window 
    {
        private AutoPal DangDangPal;
        private log DDLog;
        private string LogPath;

        public MainWindow()
        {
            InitializeComponent();
            DangDangPal = new AutoPal();
            this.DataContext = DangDangPal;
            pwdBoxADSL.Password = Settings.Default.ADSLPasswordSettings;
            LogPath = System.Environment.CurrentDirectory + "\\DDEvent.log";
            DDLog = new log();
        }
        
        private void OnMainWindow_Closed(object sender, EventArgs e)
        {
            Settings.Default.ADSLPasswordSettings = pwdBoxADSL.Password;
            Settings.Default.Save();
            CleanUp();
            Trace.TraceInformation("Rudy Trace =>Application Exited.");
        }

        

        private void OnBrowserQQ(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel(2007,2010)|*.xlsx";

            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
                DangDangPal.QQAccountFile = dlg.FileName;
        }

        private void OnBrowserTenPay(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel(2007,2010)|*.xlsx";

            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
                DangDangPal.TenpayAccountFile = dlg.FileName;
        }

        private void OnPWDChanged(object sender, RoutedEventArgs e)
        {
            DangDangPal.ADSLPassword = pwdBoxADSL.Password;
        }

        private void OnStopPalling(object sender, RoutedEventArgs e)
        {
            if (btnStop.Content.Equals("停止拍货"))
            {
                Trace.TraceInformation("Rudy Trace =>OnStopPalling: Stoppig pal...");
                DangDangPal.CancelWaitting();
                btnStop.Content = "正在取消...";
                btnStop.IsEnabled = false;
            }
            else if (btnStop.Content.Equals("返回"))
            {
                gdBeginPal.Visibility = Visibility.Visible;
                gdPalling.Visibility = Visibility.Hidden;
            }
            
        }

        private void OnBrowserBindQQAccount(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel(2007,2010)|*.xlsx";

            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
                DangDangPal.BindQQAccountFile = dlg.FileName;
        }

        private async void OnBeginBind(object sender, RoutedEventArgs e)
        {
            await PrepareEnvironmentAsync(DangDangPal.BrowserIndex);

            bool bRet = DangDangPal.SetWebDriver(DangDangPal.BrowserIndex);
            if (bRet)
                await DangDangPal.BindAllAccountAddressAsync();
            else
                Trace.TraceInformation("Rudy Trace =>OnBeginBind: Set web driver failed.");
        }

        private async void OnBeginPal(object sender, RoutedEventArgs e)
        {
            Trace.WriteLine(">>>>>>>>>>>>>>>>>>>>Rudy Trace =>OnBeginPal: Pal Start<<<<<<<<<<<<<<<<<<<<");
            gdBeginPal.Visibility = Visibility.Hidden;
            gdPalling.Visibility = Visibility.Visible;
            btnStop.Content = "停止拍货";

            Trace.WriteLine("Rudy Trace =>PrepareEnvironmentAsync: Prepare the environment...");
            await PrepareEnvironmentAsync(DangDangPal.BrowserIndex);
            Trace.WriteLine("Rudy Trace =>PrepareEnvironmentAsync: Environment ready！");

            string ProductLink = "http://product.dangdang.com/1263628906.html";
            bool bRet = DangDangPal.SetWebDriver(DangDangPal.BrowserIndex);
            try
            {
                if(bRet)
                    await DangDangPal.AutoPalProcessAsync("2508961540", "abc123456", ProductLink);
                else
                    Trace.WriteLine("Rudy Trace =>OnBeginPal: Set Web Driver Failed.");
            }
            catch (OperationCanceledException)
            {
                Trace.TraceInformation("Rudy Trace =>OnBeginPal: Pal Stopped.");
                btnStop.IsEnabled = true;
            }

            btnStop.Content = "返回";
        }

        private Task PrepareEnvironmentAsync(int BrowserIndex)
        {
            return Task.Run(() =>
            {
                if (BrowserIndex == 0)
                {
                    Process[] pBrowsers = Process.GetProcessesByName("chrome");
                    foreach (Process pBrowser in pBrowsers)
                    {
                        pBrowser.Kill();
                    }
                    Process[] pDrivers = Process.GetProcessesByName("chromedriver");
                    foreach (Process pDriver in pDrivers)
                    {
                        pDriver.Kill();
                    }
                }
                else if (BrowserIndex == 1)
                {
                    Process[] pBrowsers = Process.GetProcessesByName("iexplore");
                    foreach (Process pBrowser in pBrowsers)
                    {
                        pBrowser.Kill();
                    }
                    Process[] pDrivers = Process.GetProcessesByName("IEDriverServer");
                    foreach (Process pDriver in pDrivers)
                    {
                        pDriver.Kill();
                    }
                }
                else if (BrowserIndex == 2)
                {
                    Process[] pBrowsers = Process.GetProcessesByName("firefox");
                    foreach (Process pBrowser in pBrowsers)
                    {
                        pBrowser.Kill();
                    }
                }
                else
                {
                    Trace.TraceInformation("Rudy Trace =>Invalid Browser Type.");
                }
            });
        }

        private void CleanUp()
        {
            Trace.WriteLine("Rudy Trace =>Cleaning up the environment...");
            DangDangPal.Dispose();
            Trace.WriteLine("Rudy Trace =>Clea up done.");
        }

        private void OnReduce_Click(object sender, RoutedEventArgs e)
        {
            if (DangDangPal.SinglePalCount > 1)
                DangDangPal.SinglePalCount -= 1;
        }

        private void OnAdd_Click(object sender, RoutedEventArgs e)
        {
            DangDangPal.SinglePalCount += 1;
        }

        private void OnTest(object sender, RoutedEventArgs e)
        {
            //WebClient webclient = new WebClient();
            //webclient.getPage("www.dangdang.com");
        }
    }
}
