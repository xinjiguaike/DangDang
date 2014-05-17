using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Internal;
using OpenQA.Selenium.Support;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using System.Drawing;
using System.Windows.Automation;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Diagnostics;


namespace DangDangAutoPal.Models
{
	public struct AccountInfo
	{
		public string QQAccount;
		public string QQPassword;
		public string FullName;
		public string Mobile;
		public string Country;
		public string Province;
		public string City;
		public string Town;
		public string DetailAddress;
	}

	public class AutoPal: INotifyPropertyChanged
	{
		//Properties
		private IWebDriver driver;
		private List<AccountInfo> aAccountInfo;
        private CancellationTokenSource cts;

		private List<string> aTenAccount;
		private List<string> aTenPassword;

		private string _qqAccountFile;
		public string QQAccountFile 
		{ 
			get
			{
				return _qqAccountFile;
			} 
			set
			{
				_qqAccountFile = value;
				OnPropertyChanged("QQAccountFile");
			}
		}

		private string _tenpayAccountFile;
		public string TenpayAccountFile
		{
			get
			{
				return _tenpayAccountFile;
			}
			set
			{
				_tenpayAccountFile = value;
				OnPropertyChanged("TenpayAccountFile");
			}
		}

		private string _bindQQAccountFile;
		public string BindQQAccountFile
		{
			get
			{
				return _bindQQAccountFile;
			}
			set
			{
				_bindQQAccountFile = value;
				OnPropertyChanged("BindQQAccountFile");
			}
		
		}

		public string ProductLink { get; set; }
		public string PseudoProductLink { get; set; }
		public string Remark { get; set; }
		public string ADSLAccount { get; set; }
		public string ADSLPassword { get; set; }
		public int SinglePalCount { get; set; }
		public int TimeOut { get; set; }

		public int SuccessPalCount { get; set; }

		//Functions
		public AutoPal()
		{
			QQAccountFile = "";
			TenpayAccountFile = "";
			BindQQAccountFile = "";
			aAccountInfo = new List<AccountInfo>();
            cts = new CancellationTokenSource();
		}

		public bool SetWebDriver(string BrowserType)
		{
			if (BrowserType.Equals("Chrome"))
			{
				string ProfilePath = Environment.GetEnvironmentVariable("LocalAppData") + "\\Google\\Chrome\\User Data";
				var Options = new ChromeOptions();
				Options.AddArguments("--incognito");
				Options.AddArguments("--user-data-dir=" + ProfilePath);
				Options.AddArguments("--disable-extensions");

				driver = new ChromeDriver(Options);
				App.WindowHide(Globals.CHROME_DRIVER_TITLE);
			}
			else if (BrowserType.Equals("FireFox"))
			{
				string firefox_path = @"C:\Program Files\Mozilla Firefox\firefox.exe";
				FirefoxBinary binary = new FirefoxBinary(firefox_path);
				FirefoxProfile profile = new FirefoxProfile();
				profile.SetPreference("network.proxy.type", 0);
				driver = new FirefoxDriver(binary, profile, TimeSpan.FromSeconds(120));
			}
			else if (BrowserType.Equals("IE"))
			{
				driver = new InternetExplorerDriver();
				App.WindowHide(Globals.IE_DRIVER_TITLE);
			}
			else
			{
				Trace.TraceInformation("Rudy Trace: Invalid Browser Type.");
				return false;
			}

			driver.Manage().Window.Maximize();
			driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(60));
			driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(60));
			
			return true;
		}

		public Task CleanUpAsync()
		{
            return Task.Run(() =>
            { 
                if (driver != null)
			    {
                    cts.Cancel();
				    driver.Quit();
			    }
            });
		}

		public Task<IWebElement>  WaitForElementAsync(string el_mark, string el_flag,  int timeout = 30)
		{
            CancellationToken ct = cts.Token;
            return Task.Run(() =>
            {
                Trace.TraceInformation("Rudy Trace, Searching Element: [" + el_mark + "]");
                IWebElement ele = null;
                try
                {
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeout));
                    ele = wait.Until<IWebElement>((d) =>
                    {
                        ct.ThrowIfCancellationRequested();//
                        if (el_flag.Equals("Id"))
                            return d.FindElement(By.Id(el_mark));
                        else if (el_flag.Equals("Class"))
                            return d.FindElement(By.ClassName(el_mark));
                        else if (el_flag.Equals("Name"))
                            return d.FindElement(By.Name(el_mark));
                        else if (el_flag.Equals("XPath"))
                            return d.FindElement(By.XPath(el_mark));
                        else
                            Trace.TraceInformation("Rudy Trace, Element Flag is invalid.");
                        return null;
                    });
                }
                catch (Exception e)
                {
                    Trace.TraceInformation(e.Message.ToString() + ", Timeout to find element: [" + el_mark + "]");
                    return null;
                }
                Trace.TraceInformation("Rudy Trace, Found Element: [" + el_mark + "]");
                return ele;
            }, ct);
		}


		public Task<bool> WaitForPageAsync(string PageTitle, int Seconds)
		{
            return Task.Run(() =>
            {
                Trace.TraceInformation("Rudy Trace, WaitForPageAsync: Waitting for page [{0}]...", PageTitle);
                string defaultWindow = driver.CurrentWindowHandle;
                DateTime begins = DateTime.Now;
                TimeSpan span = DateTime.Now - begins;
                while (span.TotalSeconds < Seconds)
                {
                    foreach (string strWindow in driver.WindowHandles)
                    {
                        try
                        {
                            driver.SwitchTo().Window(strWindow);
                            if (driver.Title.Contains(PageTitle))
                            {
                                Trace.TraceInformation("Rudy Trace, WaitForPageAsync: Page [{0}] Load Succeed!", PageTitle);
                                return true;
                            }
                        }
                        catch (Exception e)
                        {
                            Trace.TraceInformation("Rudy Exception, WaitForPageAsync: " + e.Message);
                        }
                    }
                    span = DateTime.Now - begins;
                }
                //Trace.TraceInformation("Rudy Trace: Switch to default window.");
                //driver.SwitchTo().Window(defaultWindow);
                Trace.TraceInformation("Rudy Trace, WaitForPageAsync: Page [{0}] Load Time Out!", PageTitle);
                return false;
            });
		}

		public async Task<bool> SetAddressAccoutInfoAsync(string FilePath)
		{
            bool bRet = await Task.Run(() =>
            {
                if (aAccountInfo != null)
                    aAccountInfo.Clear();
                try
                {
                    Application excel = new Application();
                    excel.Visible = false;
                    Workbook wb = excel.Workbooks.Open(FilePath);
                    Worksheet ws = wb.ActiveSheet as Worksheet;
                    int nRowCount = ws.UsedRange.Cells.Rows.Count;//get the used rows count.

                    AccountInfo infoTemp;
                    infoTemp.Country = "中国";
                    for (int i = 2; i <= nRowCount; i++)
                    {
                        infoTemp.QQAccount = ((Range)ws.Cells[i, 1]).Text;
                        infoTemp.QQPassword = ((Range)ws.Cells[i, 2]).Text;
                        infoTemp.FullName = ((Range)ws.Cells[i, 3]).Text;
                        infoTemp.Mobile = ((Range)ws.Cells[i, 4]).Text;
                        infoTemp.Province = ((Range)ws.Cells[i, 5]).Text;
                        infoTemp.City = ((Range)ws.Cells[i, 6]).Text;
                        infoTemp.Town = ((Range)ws.Cells[i, 7]).Text;
                        infoTemp.DetailAddress = ((Range)ws.Cells[i, 8]).Text;
                        aAccountInfo.Add(infoTemp);
                    }
                    if (wb != null)
                        wb.Close();
                    if (excel != null)
                    {
                        excel.Quit();
                        App.KillExcel(excel);
                        excel = null;
                    }
                }
                catch (Exception e)
                {
                    System.Windows.MessageBox.Show(e.Message, "当当自动拍货");
                    Trace.TraceInformation("Rudy Exception, SetAddressAccoutInfoAsync： " + e.Source + ";" + e.Message);
                    return false;
                }

                return true;
            }).ConfigureAwait(false);
            return bRet;
		}

		public async Task<bool> LoginAsync(string Account, string Password, string ReturnPageTitle)
		{
			try
			{
				//driver.Navigate().GoToUrl(Globals.LOGIN_URL);
				var linkQQLogin = await WaitForElementAsync(Globals.QQ_LOGIN_XPATH, "XPath").ConfigureAwait(false);
				linkQQLogin.Click();
				Trace.WriteLine("Rudy Trace, LoginAsync: click QQLogin Link Successed!");

				driver.SwitchTo().Window(driver.WindowHandles.Last());
				driver.SwitchTo().Frame(Globals.QQ_FRAME_NAME);

				var tabAccoutLogin = await WaitForElementAsync(Globals.QQ_SWITCH_ACCOUNT_LOGIN_ID, "Id").ConfigureAwait(false);
				tabAccoutLogin.Click();
				Thread.Sleep(1000);

                var inputQQAccount = await WaitForElementAsync(Globals.QQ_ACCOUNT_ID, "Id").ConfigureAwait(false);
				inputQQAccount.Clear();
				inputQQAccount.SendKeys(Account);

                var inputQQPassword = await WaitForElementAsync(Globals.QQ_PASSWORD_ID, "Id").ConfigureAwait(false);
				inputQQPassword.SendKeys(Password);

                var btnQQLogin = await WaitForElementAsync(Globals.QQ_LOGINBTN_ID, "Id").ConfigureAwait(false);
				btnQQLogin.Click();
				Trace.WriteLine("Rudy Trace, LoginAsync: click QQLogin Button Successed!");
			}
			catch (Exception e)
			{
				Trace.TraceInformation("Rudy Exception, LoginAsync: " + e.Source + ";" + e.Message);
				return false;
			}

            bool bRet = await WaitForPageAsync(ReturnPageTitle, 30).ConfigureAwait(false);
			if (!bRet)
			{
				Trace.TraceInformation("Rudy Trace: Log in time out.");
				return false;
			}
			Trace.TraceInformation("Rudy Trace: Log in succeed.");
			return true;
		}

		public async Task<bool> BindingDeliverAddress(AccountInfo info)
		{
            bool bRet = await LoginAsync(info.QQAccount, info.QQPassword, "dangdang").ConfigureAwait(false);
			if (!bRet)
			{
				Trace.TraceInformation("Log in Failed!");
				return false;
			}

			try
			{
				driver.Navigate().GoToUrl(Globals.ONEKEY_BUY_URL);

                var btnAddAddress = await WaitForElementAsync(Globals.BTN_ADD_ADDRESS_CLASS, "Class", 30).ConfigureAwait(false);
				btnAddAddress.Click();

                var inputShipMan = await WaitForElementAsync(Globals.SHIP_MAN_ID, "Id").ConfigureAwait(false);
				inputShipMan.SendKeys(info.FullName);

                var selectCountry = await WaitForElementAsync(Globals.SELECT_COUNTRY_ID, "Id").ConfigureAwait(false);
                SelectElement seCountry = new SelectElement(selectCountry);
				seCountry.SelectByText(info.Country);

                var selectProvince = await WaitForElementAsync(Globals.SELECT_PROVINCE_ID, "Id").ConfigureAwait(false);
                SelectElement seProvince = new SelectElement(selectProvince);
				seProvince.SelectByText(info.Province);

                var selectCity = await WaitForElementAsync(Globals.SELECT_CITY_ID, "Id").ConfigureAwait(false);
                SelectElement seCity = new SelectElement(selectCity);
				seCity.SelectByText(info.City);

                var selectTown = await WaitForElementAsync(Globals.SELECT_CITY_ID, "Id").ConfigureAwait(false);
                SelectElement seTown = new SelectElement(selectTown);
				seTown.SelectByText(info.Town);

                var inputAddressDetail = await WaitForElementAsync(Globals.ADDRESS_DETAIL_ID, "Id").ConfigureAwait(false);
				inputAddressDetail.SendKeys(info.DetailAddress);

				//var inputZipCode = await WaitForElementAsync(Globals.ZIP_CODE_ID, "Id");
				//inputZipCode.SendKeys("310012");

                var inputMobile = await WaitForElementAsync(Globals.MOBILE_ID, "Id").ConfigureAwait(false);
				inputMobile.SendKeys(info.Mobile);

                var btnConfirmAddress = await WaitForElementAsync(Globals.CONFIRM_ADDRESS_XPATH, "XPath").ConfigureAwait(false);
				btnConfirmAddress.Click();

                var radioNormalShip = await WaitForElementAsync(Globals.RADIO_NORMALSHIP_XPATH, "XPath").ConfigureAwait(false);
				if (!radioNormalShip.Selected)
					radioNormalShip.Click();

                var btnConfirmPayment = await WaitForElementAsync(Globals.CONFIRM_PAYMENT_XPATH, "XPath").ConfigureAwait(false);
				btnConfirmPayment.Click();

                var radioNetPay = await WaitForElementAsync(Globals.RADIO_NETPAY_XPATH, "XPath").ConfigureAwait(false);
				if (!radioNetPay.Selected)
					radioNetPay.Click();

                var btnConfirmInvoice = await WaitForElementAsync(Globals.CONFIRM_INVOICE_XPATH, "XPath").ConfigureAwait(false);
				btnConfirmInvoice.Click();

                var cbxInvoice = await WaitForElementAsync(Globals.CHB_INVOICE_ID, "Id").ConfigureAwait(false);
				if (!cbxInvoice.Selected)
					cbxInvoice.Click();

                var btnCheckSubmit = await WaitForElementAsync(Globals.CHECK_SUBMIT_XPATH, "XPath").ConfigureAwait(false);
				btnCheckSubmit.Click();
			}
			catch (Exception e)
			{
				Trace.TraceInformation("Rudy Exception, BindingDeliverAddress: " + e.Message);
				return false;
			}

			return true;
		}

		public async Task BindAllAccountAddressAsync()
		{
			bool bRet = await SetAddressAccoutInfoAsync(BindQQAccountFile).ConfigureAwait(false);
			if (bRet)
			{
				Trace.TraceInformation("Rudy Trace, Set Address Account Info Success!");
				foreach (AccountInfo info in aAccountInfo)
				{
                    bRet = await BindingDeliverAddress(info).ConfigureAwait(false);
					if (bRet)
						Trace.TraceInformation("Rudy Trace, Accout[{0}]Binding Success!", info.QQAccount);
					else
						Trace.TraceInformation("Rudy Trace, Accout[{0}]Binding Failed!", info.QQAccount);
				}
			}
			else
				Trace.TraceInformation("Rudy Trace, Set Address Account Info Failed!");
		}

		public async Task<bool> PalProductAsync(string ProductLink, string Account, string Password)
		{
			try
			{
				//driver.Navigate().GoToUrl(ProductLink);
                var cbxColor = await WaitForElementAsync(Globals.COLOR_IMAGE1_XPATH, "XPath").ConfigureAwait(false);
				cbxColor.Click();

                var inputBuyNum = await WaitForElementAsync(Globals.BUY_NUM_ID, "Id").ConfigureAwait(false);
				inputBuyNum.Clear();
				inputBuyNum.SendKeys("2");

                await Task.Delay(1000).ConfigureAwait(false);//Wait for 'one key buy' button load complete.

                var btnBuyNow = await WaitForElementAsync(Globals.BUY_NOW_ID, "Id").ConfigureAwait(false);
				btnBuyNow.Click();

                var popUpBuyNow = await WaitForElementAsync(Globals.BUY_NOW_POPUP_ID, "Id").ConfigureAwait(false);
				if (popUpBuyNow != null)
				{
                    var btnConfirmOneKeyBuy = await WaitForElementAsync(Globals.BTN_CONFIRM_BUY_ID, "Id").ConfigureAwait(false);
					btnConfirmOneKeyBuy.Click();
				}
                bool bRet = await WaitForPageAsync(Globals.PAYMENT_TITLE, 30).ConfigureAwait(false);
				if (!bRet)
					return false;
			}
			catch (Exception e)
			{
                Trace.TraceInformation("Rudy Exception, PalProductAsync: " + e.Source + ";" + e.Message);
				return false;
			}
			Trace.TraceInformation("Rudy Trace: Switch to payment page succeed.");
			return true;
		}

		public async Task<bool> SelectPayPlatformAsync(string Platform)
		{
			try
			{
				var tabPayPlatform = await WaitForElementAsync(Globals.TAB_PAYMENT_PLATFORM_ID, "Id").ConfigureAwait(false);
				tabPayPlatform.Click();

				var radioTenpay = await WaitForElementAsync(Platform, "XPath").ConfigureAwait(false);
				if(!radioTenpay.Selected)
					radioTenpay.Click();

				var btnNext = await WaitForElementAsync(Globals.BTN_NEXT_ID, "Id").ConfigureAwait(false);
				btnNext.Click();

                bool bRet = await WaitForPageAsync(Globals.TENPAY_TITLE, 30).ConfigureAwait(false);
				if (!bRet)
					return false;
			}
			catch(Exception e)
			{
                Trace.TraceInformation("Rudy Exception, SelectPayPlatformAsync: " + e.Source + ";" + e.Message);
				return false;
			}
			Trace.TraceInformation("Rudy Trace: Switch to Tenpay page succeed.");
			return true;
		}

        public async Task<bool> TenpayAsync(string TenpayUser, string TenpayPass)
		{
			try
			{ 
				var inputTenpayUser = await WaitForElementAsync(Globals.TENPAY_USERNAME_ID, "Id").ConfigureAwait(false);
				inputTenpayUser.Clear();
				inputTenpayUser.SendKeys(TenpayUser);
                var inputTenpayPass = await WaitForElementAsync(Globals.TENPAY_PASSWORD_ID, "Id").ConfigureAwait(false);
				inputTenpayPass.SendKeys(TenpayPass);
                var btnLogin = await WaitForElementAsync(Globals.TENPAY_LOGIN_ID, "Id").ConfigureAwait(false);
				btnLogin.Click();

				//driver.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(60));
                var radioBalance = await WaitForElementAsync(Globals.RADIO_BALANCEPAY_CLASS, "Class").ConfigureAwait(false);
				if (!radioBalance.Selected && radioBalance.Displayed)
					radioBalance.Click();

                var ctrlPassword = await WaitForElementAsync(Globals.CTRL_TENPWD_LABEL_XPATH, "XPath").ConfigureAwait(false);

				/*var btnConfirmToPay = await WaitForElementAsync(Globals.CONFIRM_TO_PAY_ID, "Id");
				btnConfirmToPay.Click();*/

				System.Drawing.Point ctrlPassPosition = ctrlPassword.Location;
				OpenQA.Selenium.Interactions.Actions ctrlAction = new OpenQA.Selenium.Interactions.Actions(driver);
				Trace.TraceInformation("Rudy Trace, ctrlPassPosition.X = " + ctrlPassPosition.X.ToString());
				Trace.TraceInformation("Rudy Trace, ctrlPassPosition.Y = " + ctrlPassPosition.Y.ToString());
				//ctrlAction.MoveToElement(ctrlPassword, 10, 30);
				//ctrlAction.Click();
                await Task.Delay(10000).ConfigureAwait(false);

				ctrlAction.SendKeys("12345");
				//System.Windows.MessageBox.Show("TenPay Succeed!", "Rudy");
				 
				 /*
				AutomationElement ctrlPassword = FindChildElementByClass("Edit", "Chrome_WidgetWin_1");
				if (ctrlPassword == null)
				{
					Trace.TraceInformation("Rudy Trace, ctrlPassword is null.");
					return false;
				}
				var patternValue = (ValuePattern)ctrlPassword.GetCurrentPattern(ValuePattern.Pattern);
				if (patternValue != null)
					patternValue.SetValue(TenpayPass);
				else
				{
					Trace.TraceInformation("Rudy Trace, patternValue is null.");
					return false;
				}*/

				Trace.TraceInformation("Rudy Trace, Tenpay Succeed.");
			}
			catch(Exception e)
			{
                Trace.TraceInformation("Rudy Exception, TenpayAsync: " + e.Source + ";" + e.Message);
				return false;
			}

			return true;
		}

		public async Task<bool> AutoPalProcessAsync(string Account, string Password, string ProductLink)
		{
            await Task.Run(() =>
                { 
                    driver.Navigate().GoToUrl(ProductLink);
                }).ConfigureAwait(false); 

			string PageTitle = driver.Title;

            var linkLogin = await WaitForElementAsync(Globals.LOGIN_LINK_CLASS, "Class").ConfigureAwait(false);
			linkLogin.Click();
			Trace.WriteLine("Rudy Trace, AutoPalProcessAsync: click Login Link Successed!");

            bool bRet = await WaitForPageAsync(Globals.LOGIN_PAGE_TITLE, 30).ConfigureAwait(false);
            if (!bRet)
				return false;

            bRet = await LoginAsync(Account, Password, PageTitle).ConfigureAwait(false);
            if (!bRet)
			{
				Trace.TraceInformation("Rudy Trace, AutoPalProcessAsync: Log in Failed!");
				return false;
			}

            bRet = await PalProductAsync(ProductLink, Account, Password).ConfigureAwait(false);
            if (!bRet)
			{
				Trace.TraceInformation("Rudy Trace, AutoPalProcessAsync: Pal Product Failed!");
				return false;
			}

            bRet = await SelectPayPlatformAsync(Globals.RADIO_TENPAY_XPATH).ConfigureAwait(false);
            if (!bRet)
			{
                Trace.TraceInformation("Rudy Trace, AutoPalProcessAsync: Select Pay Platform Failed!");
				return false;
			}

			string TenpayUserName = "672045573";
			string TenpayPWD = "mbi@88820";
            bRet = await TenpayAsync(TenpayUserName, TenpayPWD).ConfigureAwait(false);
            if (!bRet)
			{
				Trace.TraceInformation("AutoPalProcessAsync: TenPay Failed!");
				return false;
			}

			return true;
		}

		private AutomationElement FindChildElementByClass(string ClassName, string BrowserWndClassName)
		{
			try
			{
				Condition propCondition = new PropertyCondition(AutomationElement.ClassNameProperty, BrowserWndClassName);
				AutomationElement rootElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, propCondition);
				if (rootElement == null)
				{
					Trace.TraceInformation("Rudy Trace, rootElement is null.");
					return null;
				}

				propCondition = new PropertyCondition(AutomationElement.ClassNameProperty, "WrapperNativeWindowClass");
				AutomationElement wrapperElement = rootElement.FindFirst(TreeScope.Children, propCondition);
				if (wrapperElement == null)
				{
					Trace.TraceInformation("Rudy Trace, wrapperElement is null.");
					return null;
				}
				

				propCondition = new PropertyCondition(AutomationElement.ClassNameProperty, ClassName);
				AutomationElement childElement = wrapperElement.FindFirst(TreeScope.Descendants, propCondition);
				if(childElement == null)
				{
					Trace.TraceInformation("Rudy Trace, childElement is null.");
					return null;
				}
				return childElement;
			}
			catch(Exception e)
			{
                Trace.TraceInformation("Rudy Exception, FindChildElementByClass: " + e.Source + ";" + e.Message);
				return null;
			}
			
		}

		public event PropertyChangedEventHandler PropertyChanged;
		public void OnPropertyChanged(string propertyName)
		{
			if (PropertyChanged != null)
				PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
		}
	}
}
