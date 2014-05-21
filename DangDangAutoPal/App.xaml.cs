﻿using System;
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
        private static readonly int MOUSEEVENTF_LEFTDOWN = 0x0002;//模拟鼠标移动
        private static readonly int MOUSEEVENTF_MOVE = 0x0001;//模拟鼠标左键按下
        private static readonly int MOUSEEVENTF_LEFTUP = 0x0004;//模拟鼠标左键抬起
        private static readonly int MOUSEEVENTF_ABSOLUTE = 0x8000;//鼠标绝对位置
        //private readonly int MOUSEEVENTF_RIGHTDOWN = 0x0008; //模拟鼠标右键按下 
        //private readonly int MOUSEEVENTF_RIGHTUP = 0x0010; //模拟鼠标右键抬起 
        //private readonly int MOUSEEVENTF_MIDDLEDOWN = 0x0020; //模拟鼠标中键按下 
        //private readonly int MOUSEEVENTF_MIDDLEUP = 0x0040;// 模拟鼠标中键抬起 

        [DllImport("user32.dll", EntryPoint = "ShowWindow", SetLastError = true)]
        private static extern bool ShowWindow(IntPtr hWnd, uint nCmdShow);
        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);

        [DllImport("user32")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        public static void ClickLeft(int x, int y)
        {
            //绝对位置
            mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, x * 65535 / 1600, y * 65535 / 900, 0, 0);//移动到需要点击的位置
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_ABSOLUTE, x * 65535 / 1600, y * 65535 / 900, 0, 0);//点击
            mouse_event(MOUSEEVENTF_LEFTUP | MOUSEEVENTF_ABSOLUTE, x * 65535 / 1600, y * 65535 / 900, 0, 0);//抬

        }


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
