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


namespace DangDangAutoPal
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window 
	{
		private AutoPal DangDangPal;
		public MainWindow()
		{
			InitializeComponent();
			DangDangPal = new AutoPal();
			this.DataContext = DangDangPal;
			pwdBoxADSL.Password = Settings.Default.ADSLPasswordSettings;
		}
		
		private async void MainWindow_Closed(object sender, EventArgs e)
		{
			Settings.Default.ADSLPasswordSettings = pwdBoxADSL.Password;
			Settings.Default.Save(); 
			await CleanUpAsync();
			Trace.TraceInformation("Rudy Trace, Application Exited.");
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
			Trace.TraceInformation("Rudy Trace, OnStopPalling: Pal Stopped.");
			gdBeginPal.Visibility = Visibility.Visible;
			gdPalling.Visibility = Visibility.Hidden;
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
			await PrepareEnvironmentAsync("Chrome");

			bool bRet = DangDangPal.SetWebDriver("Chrome");
			if (bRet)
				await DangDangPal.BindAllAccountAddressAsync();
			else
				Trace.TraceInformation("Rudy Trace, OnBeginBind: Set Web Driver Failed.");
		}

		private async void OnBeginPal(object sender, RoutedEventArgs e)
		{
			Trace.WriteLine("Rudy Trace, OnBeginPal: Begin Pal! ");
			gdBeginPal.Visibility = Visibility.Hidden;
			gdPalling.Visibility = Visibility.Visible;

			await PrepareEnvironmentAsync("Chrome");

			string ProductLink = "http://product.dangdang.com/1263628906.html";
			bool bRet = DangDangPal.SetWebDriver("Chrome");
			if (bRet)
                await DangDangPal.AutoPalProcessAsync("672045573", "mbi@88820", ProductLink);
			else
				Trace.WriteLine("Rudy Trace, OnBeginPal: Set Web Driver Failed.");
		}

		private Task PrepareEnvironmentAsync(string BrowserType)
		{
			Trace.WriteLine("Rudy Trace, PrepareEnvironmentAsync: Begin Prepare Environment！");
            return Task.Run(() =>
            {
                if (BrowserType.Equals("Chrome"))
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
            });
		}

		private async Task CleanUpAsync()
		{
			Trace.WriteLine("Rudy Trace, Cleaning up the Environment...");
            if (DangDangPal != null)
                await DangDangPal.CleanUpAsync().ConfigureAwait(false);
            Trace.WriteLine("Rudy Trace, Clea up done.");
		}

        private async void OnTest_Click(object sender, RoutedEventArgs e)
        {
            await PrepareEnvironmentAsync("Chrome");
            bool bRet = DangDangPal.SetWebDriver("Chrome");
            if (bRet)
                await DangDangPal.WaitForElementAsync("rudy", "Class", 10);
            else
                Trace.WriteLine("Rudy Trace, OnTest_Click: Set Web Driver Failed.");
        }
	}
}
