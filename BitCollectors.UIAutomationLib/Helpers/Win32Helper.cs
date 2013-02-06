using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using BitCollectors.UIAutomationLib.Entities;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace BitCollectors.UIAutomationLib.Helpers
{
    public class Win32Helper
    {
        #region Enums
        public enum SystemMetric
        {
            SM_CXSCREEN = 0,
            SM_CYSCREEN = 1,
        }

        public enum VirtualKeys : ushort
        {
            SHIFT = 0x10,
            CONTROL = 0x11,
            MENU = 0x12,
            ESCAPE = 0x1B,
            BACK = 0x08,
            TAB = 0x09,
            RETURN = 0x0D,
            PRIOR = 0x21,
            NEXT = 0x22,
            END = 0x23,
            HOME = 0x24,
            LEFT = 0x25,
            UP = 0x26,
            RIGHT = 0x27,
            DOWN = 0x28,
            SELECT = 0x29,
            PRINT = 0x2A,
            EXECUTE = 0x2B,
            SNAPSHOT = 0x2C,
            INSERT = 0x2D,
            DELETE = 0x2E,
            HELP = 0x2F,
            NUMPAD0 = 0x60,
            NUMPAD1 = 0x61,
            NUMPAD2 = 0x62,
            NUMPAD3 = 0x63,
            NUMPAD4 = 0x64,
            NUMPAD5 = 0x65,
            NUMPAD6 = 0x66,
            NUMPAD7 = 0x67,
            NUMPAD8 = 0x68,
            NUMPAD9 = 0x69,
            MULTIPLY = 0x6A,
            ADD = 0x6B,
            SEPARATOR = 0x6C,
            SUBTRACT = 0x6D,
            DECIMAL = 0x6E,
            DIVIDE = 0x6F,
            F1 = 0x70,
            F2 = 0x71,
            F3 = 0x72,
            F4 = 0x73,
            F5 = 0x74,
            F6 = 0x75,
            F7 = 0x76,
            F8 = 0x77,
            F9 = 0x78,
            F10 = 0x79,
            F11 = 0x7A,
            F12 = 0x7B,
            OEM_1 = 0xBA,   // ',:' for US
            OEM_PLUS = 0xBB,   // '+' any country
            OEM_COMMA = 0xBC,   // ',' any country
            OEM_MINUS = 0xBD,   // '-' any country
            OEM_PERIOD = 0xBE,   // '.' any country
            OEM_2 = 0xBF,   // '/?' for US
            OEM_3 = 0xC0,   // '`~' for US
            MEDIA_NEXT_TRACK = 0xB0,
            MEDIA_PREV_TRACK = 0xB1,
            MEDIA_STOP = 0xB2,
            MEDIA_PLAY_PAUSE = 0xB3,
            LWIN = 0x5B,
            RWIN = 0x5C
        }
        #endregion

        #region User32 Externs
        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, ref INPUT pInputs, int cbSize);

        [DllImport("user32.dll")]
        private static extern uint MapVirtualKey(uint uCode, uint uMapType);

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern int SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowRect(IntPtr hWnd, ref RECT rect);

        [DllImport("user32.dll")]
        public static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        //[DllImport("user32.dll")]
        //[return: MarshalAs(UnmanagedType.Bool)]
        //private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        #endregion

        #region GDI32 Externs
        [DllImport("gdi32.dll")]
        public static extern bool BitBlt(IntPtr hObject, int nXDest, int nYDest, int nWidth, int nHeight, IntPtr hObjectSource, int nXSrc, int nYSrc, int dwRop);

        [DllImport("gdi32.dll")]
        public static extern IntPtr CreateCompatibleBitmap(IntPtr hDC, int nWidth, int nHeight);

        [DllImport("gdi32.dll")]
        public static extern IntPtr CreateCompatibleDC(IntPtr hDC);

        [DllImport("gdi32.dll")]
        public static extern bool DeleteDC(IntPtr hDC);

        [DllImport("gdi32.dll")]
        public static extern bool DeleteObject(IntPtr hObject);

        [DllImport("gdi32.dll")]
        public static extern IntPtr SelectObject(IntPtr hDC, IntPtr hObject);

        [DllImport("user32.dll")]
        static extern int GetSystemMetrics(SystemMetric smIndex);
        #endregion

        #region Structs
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        [StructLayout(LayoutKind.Explicit)]
        struct INPUT
        {
            [FieldOffset(0)]
            public int type;
            [FieldOffset(4)]
            public MOUSEINPUT mi;
            [FieldOffset(4)]
            public KEYBDINPUT ki;
            [FieldOffset(4)]
            public HARDWAREINPUT hi;
        }
        #endregion

        #region Constants
        const int INPUT_MOUSE = 0;
        const int INPUT_KEYBOARD = 1;
        const int INPUT_HARDWARE = 2;
        const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
        const uint KEYEVENTF_KEYUP = 0x0002;
        const uint KEYEVENTF_UNICODE = 0x0004;
        const uint KEYEVENTF_SCANCODE = 0x0008;
        const uint XBUTTON1 = 0x0001;
        const uint XBUTTON2 = 0x0002;
        const uint MOUSEEVENTF_MOVE = 0x0001;
        const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        const uint MOUSEEVENTF_LEFTUP = 0x0004;
        const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        const uint MOUSEEVENTF_RIGHTUP = 0x0010;
        const uint MOUSEEVENTF_MIDDLEDOWN = 0x0020;
        const uint MOUSEEVENTF_MIDDLEUP = 0x0040;
        const uint MOUSEEVENTF_XDOWN = 0x0080;
        const uint MOUSEEVENTF_XUP = 0x0100;
        const uint MOUSEEVENTF_WHEEL = 0x0800;
        const uint MOUSEEVENTF_VIRTUALDESK = 0x4000;
        const uint MOUSEEVENTF_ABSOLUTE = 0x8000;
        #endregion

        #region Methods
        /// <summary>
        /// Converts character to hex code
        /// </summary>
        /// <param name="c">Character value to convert</param>
        /// <returns>Converted hex value</returns>
        public static short GetCharacterValue(char c)
        {
            //http://msdn.microsoft.com/en-us/library/dd375731(v=VS.85).aspx
            if (c >= 'a' && c <= 'z') return (short)(c - 'a' + 0x41);
            else if (c >= '0' && c <= '9') return (short)(c - '0' + 0x30);
            else if (c == '-') return 0x6D; // Note it's NOT 0x2D as in ASCII code!
            else if (c == ' ') return 0x20;
            else return 0; // default
        }

        /// <summary>
        /// Execute the automation process based on the supplied business object with the configuration information
        /// </summary>
        /// <param name="uiAutomation">The business object which contains configuration data and the steps to execute</param>
        public void ExecuteAutomation(UIAutomation uiAutomation)
        {
            IntPtr windowHandle = IntPtr.Zero;

            if (!string.IsNullOrEmpty(uiAutomation.FormHandleTitleBar))
            {
                windowHandle = FindWindow(null, uiAutomation.FormHandleTitleBar);
            }
            else if (!string.IsNullOrEmpty(uiAutomation.FormHandleClassName))
            {
                windowHandle = FindWindow(uiAutomation.FormHandleClassName, null);
            }
            else if (!string.IsNullOrEmpty(uiAutomation.FormHandleTitleBarRegex))
            {
                Regex regex = new Regex(uiAutomation.FormHandleTitleBarRegex);

                foreach (Process process in Process.GetProcesses())
                {
                    if (regex.IsMatch(process.MainWindowTitle))
                    {
                        windowHandle = FindWindow(null, process.MainWindowTitle);
                        
                        break;
                    }
                }
            }
            else
            {
                throw new Exception("No window handle defined");
            }


            if (!windowHandle.Equals(IntPtr.Zero))
            {
                ShowWindow(windowHandle, SW_RESTORE); 
                SetFocus(windowHandle);
                SetForegroundWindow(windowHandle);

                if (IsValidWindow(uiAutomation, windowHandle))
                {
                    foreach (InputBatch batch in uiAutomation.InputBatches)
                    {
                        if (batch.Timeout > 0)
                        {
                            Thread.Sleep(batch.Timeout);
                        }

                        SimulateKeystrokes(windowHandle, batch);
                    }
                }
            }
        }

        private const int SW_RESTORE = 9;

        private bool IsValidWindow(UIAutomation uiAutomation, IntPtr windowHandle)
        {
            if (uiAutomation.VerifyScreenOptions.VerifyScreenType == VerifyScreenTypes.PixelColor)
            {
                if (!windowHandle.Equals(IntPtr.Zero))
                {
                    Image windowImage = CaptureWindow(windowHandle);
                    Bitmap bmp = new Bitmap(windowImage);

                    Color pixelColor =
                        bmp.GetPixel(uiAutomation.VerifyScreenOptions.VerifyScreenPixelLocation.X,
                                     uiAutomation.VerifyScreenOptions.VerifyScreenPixelLocation.Y);

                    Color testColor = uiAutomation.VerifyScreenOptions.VerifyScreenPixelColor;

                    return (pixelColor == testColor);
                }

                return false;
            }

            return true;
        }

        private void SimulateMouseMove(int x, int y)
        {
            INPUT[] inp = new INPUT[2];
            inp[0].type = INPUT_MOUSE;
            inp[0].mi = CreateMouseInput(0, 0, 0, 0, MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE);
            inp[1].type = INPUT_MOUSE;
            inp[1].mi = CreateMouseInput(x, y, 0, 0, MOUSEEVENTF_MOVE);
            SendInput((uint)inp.Length, inp, Marshal.SizeOf(inp[0].GetType()));
        }

        private void SimulateMouseClick()
        {
            INPUT[] inp = new INPUT[2];
            inp[0].type = INPUT_MOUSE;
            inp[0].mi = CreateMouseInput(0, 0, 0, 0, MOUSEEVENTF_LEFTDOWN);
            inp[1].type = INPUT_MOUSE;
            inp[1].mi = CreateMouseInput(0, 0, 0, 0, MOUSEEVENTF_LEFTUP);
            SendInput((uint)inp.Length, inp, Marshal.SizeOf(inp[0].GetType()));
        }

        private void SimulateKeystrokes(IntPtr windowHandle, InputBatch batch)
        {
            List<INPUT> inputList = new List<INPUT>();
            //INPUT[] inp = new INPUT[batch.InputActions.Count];

            for (int i = 0; i < batch.InputActions.Count; i++)
            {
                switch (batch.InputActions[i].InputAction)
                {
                    case InputActionTypes.KeyStroke:
                        KeyStroke keyStroke = batch.InputActions[i] as KeyStroke;
                        INPUT kbInput = new INPUT();

                        //inp[i].type = INPUT_KEYBOARD;
                        //if (keyStroke.Type == KeyStrokeTypes.Down)
                        //{
                        //    inp[i].ki = CreateKeyboardInput(keyStroke.Value, 0);
                        //}
                        //else
                        //{
                        //    inp[i].ki = CreateKeyboardInput(keyStroke.Value, KEYEVENTF_KEYUP);
                        //}


                        kbInput.type = INPUT_KEYBOARD;
                        if (keyStroke.Type == KeyStrokeTypes.Down)
                        {
                            kbInput.ki = CreateKeyboardInput(keyStroke.Value, 0);
                        }
                        else
                        {
                            kbInput.ki = CreateKeyboardInput(keyStroke.Value, KEYEVENTF_KEYUP);
                        }

                        inputList.Add(kbInput);
                        break;

                    case InputActionTypes.MouseClick:
                        MouseClick mouseClick = batch.InputActions[i] as MouseClick;
                        INPUT mouseInput = new INPUT();

                        int x = (mouseClick.Coordinates.X * 65536) / GetSystemMetrics(SystemMetric.SM_CXSCREEN);
                        int y = (mouseClick.Coordinates.Y * 65536) / GetSystemMetrics(SystemMetric.SM_CYSCREEN);

                        if (mouseClick.RelativeTo.ToLower() != "screen")
                        {
                            RECT windowRect = new RECT();
                            GetWindowRect(windowHandle, ref windowRect);

                            x = ((mouseClick.Coordinates.X + windowRect.Left) * 65536) / GetSystemMetrics(SystemMetric.SM_CXSCREEN);
                            y = ((mouseClick.Coordinates.Y + windowRect.Top) * 65536) / GetSystemMetrics(SystemMetric.SM_CYSCREEN);
                        }

                        mouseInput.type = INPUT_MOUSE;

                        switch (mouseClick.Type)
                        {
                            case MouseClickTypes.Move:
                                mouseInput.mi = CreateMouseInput(x, y, 0, 0, MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE);
                                break;

                            case MouseClickTypes.Down:
                                mouseInput.mi = CreateMouseInput(x, y, 0, 0, MOUSEEVENTF_LEFTDOWN);
                                break;

                            case MouseClickTypes.Up:
                                mouseInput.mi = CreateMouseInput(x, y, 0, 0, MOUSEEVENTF_LEFTUP);
                                break;
                        }


                        SendInput((uint)1, ref mouseInput, Marshal.SizeOf(new INPUT()));
                        //inputList.Add(mouseInput);
                        break;

                }
            }

            //SendInput((uint)inp.Length, inp, Marshal.SizeOf(inp[0].GetType()));

            if (inputList.Count > 0)
            {
                INPUT[] inp = inputList.ToArray();
                SendInput((uint)inp.Length, inp, Marshal.SizeOf(new INPUT()));
            }
        }

        private MOUSEINPUT CreateMouseInput(int x, int y, uint data, uint t, uint flag)
        {
            MOUSEINPUT mi = new MOUSEINPUT();

            mi.dx = x;
            mi.dy = y;
            mi.mouseData = data;
            //mi.time = t;
            mi.dwFlags = flag;

            return mi;
        }

        private KEYBDINPUT CreateKeyboardInput(short wVK, uint flag)
        {
            KEYBDINPUT i = new KEYBDINPUT();

            i.wVk = (ushort)wVK;
            i.wScan = (ushort)MapVirtualKey((uint)wVK, 0);
            i.time = 0;
            i.dwExtraInfo = IntPtr.Zero;
            i.dwFlags = flag;

            return i;
        }
        #endregion












        public const int SRCCOPY = 13369376;

        public static Image CaptureWindow(IntPtr handle)
        {
            IntPtr hdcSrc = GetWindowDC(handle);

            RECT windowRect = new RECT();
            GetWindowRect(handle, ref windowRect);

            int width = windowRect.Right - windowRect.Left;
            int height = windowRect.Bottom - windowRect.Top;

            IntPtr hdcDest = CreateCompatibleDC(hdcSrc);
            IntPtr hBitmap = CreateCompatibleBitmap(hdcSrc, width, height);

            IntPtr hOld = SelectObject(hdcDest, hBitmap);
            BitBlt(hdcDest, 0, 0, width, height, hdcSrc, 0, 0, SRCCOPY);
            SelectObject(hdcDest, hOld);
            DeleteDC(hdcDest);
            ReleaseDC(handle, hdcSrc);

            Image image = Image.FromHbitmap(hBitmap);
            DeleteObject(hBitmap);

            SaveScreenshotFile(image, "C:\\TEMP\\Screenshot1.tif");

            return image;
        }

        private static void SaveScreenshotFile(Image image, string filename)
        {
            try
            {
                if (File.Exists(filename))
                    File.Delete(filename);
                image.Save(filename, ImageFormat.Tiff);
            }
            catch (Exception)
            {
            }
        }
    }
}
