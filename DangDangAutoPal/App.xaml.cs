using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using OpenQA.Selenium;

namespace DangDangAutoPal
{
	/// <summary>
	/// Interaction logic for App.xaml
	/// </summary>
	public partial class App : Application
	{
		[DllImport("user32.dll", EntryPoint = "ShowWindow", SetLastError = true)]
		private static extern bool ShowWindow(IntPtr hWnd, uint nCmdShow);
		[DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
		private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		private static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
		/// <summary>
		/// Hiding Window
		/// </summary>
		/// <param name="consoleTitle">Title of the Window</param>
		public static void WindowHide(string consoleTitle)
		{
			IntPtr a = FindWindow("ConsoleWindowClass", consoleTitle);
			if (a != IntPtr.Zero)
				ShowWindow(a, 0);//hide the window
			else
				throw new Exception("can't hide console window");
		}

		public static void KillExcel(Microsoft.Office.Interop.Excel.Application excel)
		{
			IntPtr t = new IntPtr(excel.Hwnd);
			int k = 0;
			GetWindowThreadProcessId(t, out k);
			System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);
			p.Kill();
		}
	}
}
