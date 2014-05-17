using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DangDangAutoPal.Models
{
	public class Globals
	{
		public const string CHROME_DRIVER_TITLE = "C:\\Windows\\system32\\chromedriver.exe";
		public const string IE_DRIVER_TITLE = "C:\\Windows\\system32\\internetexplorerdriver.exe";

		#region======================Page Title======================
		public const string LOGGEDIN_TITLE = "当当网—网上购物中心";
		public const string PAYMENT_TITLE = "当当网--支付中心";
		public const string TENPAY_TITLE = "财付通 - 支付中心";
        public const string LOGIN_PAGE_TITLE = "登录 - 当当网";
		#endregion

		#region======================Page Url========================
		public const string LOGIN_URL = "https://login.dangdang.com/signin.aspx";
		public const string ONEKEY_BUY_URL = "http://customer.dangdang.com/onekey_buy/info.php";
		#endregion

		#region======================Page Elements===================
		//Log in information
		public const string LOGIN_LINK_CLASS = "login_link";
		public const string QQ_LOGIN_XPATH = "//a[@title='QQ']";
		public const string QQ_FRAME_NAME = "ptlogin_iframe";
		public const string QQ_ACCOUNT_ID = "u";
		public const string QQ_PASSWORD_ID = "p";
		public const string QQ_LOGINBTN_ID = "login_button";
		public const string QQ_SWITCH_ACCOUNT_LOGIN_ID = "switcher_plogin";

		public const string BTN_ADD_ADDRESS_CLASS = "add_address_btn";

		//Delivery Address Infomation
		public const string SHIP_MAN_ID = "ship_man";
		public const string SELECT_COUNTRY_ID = "country_id";
		public const string SELECT_PROVINCE_ID = "province_id";
		public const string SELECT_CITY_ID = "city_id";
		public const string SELECT_TOWN_ID = "town_id";
		public const string ADDRESS_DETAIL_ID = "addr_detail";
		public const string ZIP_CODE_ID = "ship_zip";
		public const string MOBILE_ID = "ship_mb";
		public const string CONFIRM_ADDRESS_XPATH = "//a[@href='javascript:show_shipment();']";
		public const string RADIO_NORMALSHIP_XPATH = "//input[@name='ship_type' and @value='1']";
		public const string CONFIRM_PAYMENT_XPATH = "//a[@href='javascript:show_payment();']";

		public const string RADIO_NETPAY_XPATH = "//input[@name='pay_id' and @value='-1']";
		public const string CONFIRM_INVOICE_XPATH = "//a[@href='javascript:show_invoice();']";

		public const string CHB_INVOICE_ID = "no_need_invoice";
		public const string CHECK_SUBMIT_XPATH = "//a[@onclick='javascript:check_submit();return false;']";

		//Product Page
		public const string COLOR_IMAGE1_XPATH = "//a[@id='cl_0']/img";
		public const string COLOR_IMAGE2_XPATH = "//a[@id='cl_1']/img";
		public const string COLOR_IMAGE3_XPATH = "//a[@id='cl_2']/img";
		public const string COLOR_IMAGE4_XPATH = "//a[@id='cl_3']/img";
		public const string BUY_NUM_ID = "buy_num";
		public const string BUY_NOW_ID = "buy_now_button";
		public const string ADD_TO_CART_ID = "part_buy_button";
		public const string BUY_NOW_XPATH = "//div[@class='btn_p']/a[@id='buy_now_button']";
		public const string BUY_NOW_POPUP_ID = "div_onekey_select_pop";
		public const string BTN_CONFIRM_BUY_ID = "onekey_select_pop_confirm";

		//Payment page
		public const string TAB_PAYMENT_PLATFORM_ID = "go_tab3";
		public const string RADIO_TENPAY_XPATH = "//p[@bid='44']/input";
		public const string BTN_NEXT_ID = "A4";

		//Tenpay Login Page
		public const string TENPAY_LOGIN_ERR_MSG_ID = "login_err_msg";
		public const string TENPAY_USERNAME_ID = "login_uin";
		public const string TENPAY_PASSWORD_ID = "login_pwd";
		public const string TENPAY_LOGIN_ID = "btn_login";

		//Tenpay Payment Page
		public const string RADIO_BALANCEPAY_CLASS = "ctrl-radio";
		public const string CTRL_TENPWD_LABEL_XPATH = "//div[@id='paypwd_line']/label";
		public const string CONFIRM_TO_PAY_ID = "btn_pay_submit";
		#endregion
	}
}
