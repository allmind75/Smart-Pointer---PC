using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PPt = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Diagnostics;

//bluetooh
using System.Threading;
using System.Net.Sockets;
using InTheHand;
using InTheHand.Net.Bluetooth;
using InTheHand.Net.Ports;
using InTheHand.Net.Sockets;
using InTheHand.Windows.Forms;
using InTheHand.Net.Bluetooth.AttributeIds;
using System.IO;

namespace PPT_TEST
{
    public partial class MainForm1 : Form
    {

        ////마우스 후킹 user32.dll 에서 필요한 메서드를 가져온다   :: using System.Runtime.InteropServices; 사용
        [DllImport("user32.dll", EntryPoint = "SetCursorPos")]
        internal extern static Int32 SetCursorPos(Int32 x, Int32 y);
        [DllImport("user32.dll", EntryPoint = "mouse_event")]
        internal extern static Int32 mouse_event(uint dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
        ////
        //마우스 이벤트
        [DllImport("user32.dll")]
        static extern void mouse_event(uint dwFlags, uint dx, uint dy, int dwData, int dwExtraInfo);

        //마우스 Cursor
        [DllImport("user32.dll")]
        public static extern int SetSystemCursor(int hcur, int id);
        [DllImport("user32.dll")]
        public static extern int LoadCursorFromFile(string lpFileName);
        [DllImport("user32.dll")]
        public static extern int SetCursor(int hCursor);

        //키보드 후킹
        [DllImport("user32.dll")]
        static extern uint keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtratinfo);

        // Define PowerPoint Application object
        PPt.Application pptApplication;
        // Define Presentation object
        PPt.Presentation presentation;
        // Define Slide collection
        PPt.Slides slides;
        PPt.Slide slide;

        // Slide count
        int slidescount;
        // slide index
        int slideIndex;

        int x_prev = 900, y_prev = 500;

        /// <summary>
        /// 블루투스 연동
        /// </summary>
        List<string> items;
        BluetoothListener blueListener;

        public MainForm1()
        {
            InitializeComponent();

            //Bluetooth Start
            if (serverStarted)
            {
                updateUI("Server already started silly sausage!");
            }
            else
            {
                connectAsServer();
            }

            // Set Control button disable
            this.btnFirst.Enabled = false;
            this.btnNext.Enabled = false;
            this.btnPrev.Enabled = false;
            this.btnLast.Enabled = false;

            items = new List<string>();

            //트레이
            this.FormClosing += MainForm1_FormClosing;
            this.notifyIcon1.DoubleClick += notifyIcon1_DoubleClick;
            this.ExitToolStripMenuItem.Click += ExitToolStripMenuItem_Click;

        }

        private void btnCheck_Click_Click(object sender, EventArgs e)
        {
            try
            {
                // Get Running PowerPoint Application object 
                pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPt.Application;

                // Get PowerPoint application successfully, then set control button enable 
                this.btnFirst.Enabled = true;
                this.btnNext.Enabled = true;
                this.btnPrev.Enabled = true;
                this.btnLast.Enabled = true;
            }
            catch
            {
                MessageBox.Show("Please Run PowerPoint Firstly", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
            }
            if (pptApplication != null)
            {
                // Get Presentation Object 
                presentation = pptApplication.ActivePresentation;
                // Get Slide collection object 
                slides = presentation.Slides;
                // Get Slide count 
                slidescount = slides.Count;
                // Get current selected slide  

                try
                {
                    // Get selected slide object in normal view 
                    slide = slides[pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];
                }
                catch
                {
                    // Get selected slide object in reading view 
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
        }

        private void btnFirst_Click_Click(object sender, EventArgs e)
        {
            try
            {
                // Call Select method to select first slide in normal view
                slides[1].Select();
                slide = slides[1];
            }
            catch
            {
                // Transform to first page in reading view
                pptApplication.SlideShowWindows[1].View.First();
                slide = pptApplication.SlideShowWindows[1].View.Slide;
            }
        }

        private void btnNext_Click_Click(object sender, EventArgs e)
        {
            slideIndex = slide.SlideIndex + 1;
            if (slideIndex > slidescount)
            {
                //MessageBox.Show("It is already last page");
            }
            else
            {
                try
                {
                    slide = slides[slideIndex];
                    slides[slideIndex].Select();
                }
                catch
                {
                    pptApplication.SlideShowWindows[1].View.Next();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
        }

        private void btnPrev_Click_Click(object sender, EventArgs e)
        {
            slideIndex = slide.SlideIndex - 1;
            if (slideIndex >= 1)
            {
                try
                {
                    slide = slides[slideIndex];
                    slides[slideIndex].Select();
                }
                catch
                {
                    pptApplication.SlideShowWindows[1].View.Previous();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
            else
            {
                //MessageBox.Show("It is already Fist Page");

            }
        }

        private void btnlast_Click_Click(object sender, EventArgs e)
        {
            try
            {
                slides[slidescount].Select();
                slide = slides[slidescount];
            }
            catch
            {
                pptApplication.SlideShowWindows[1].View.Last();
                slide = pptApplication.SlideShowWindows[1].View.Slide;
            }
        }


        // mouseFlasg 상수
        public enum MouseFlags
        {
            LEFTDOWN = 0x00000002,
            LEFTUP = 0x00000004,
            LEFTCLICK = 0x203,
            MIDDLEDOWN = 0x00000020,
            MIDDLEUP = 0x00000040,
            MOVE = 0x00000001,
            ABSOLUTE = 0x00008000,
            RIGHTDOWN = 0x00000008,
            RIGHTUP = 0x00000010,
            RIGHTCLICK = 0x206,
            MOUSE_WHEEL = 0x00000800
        }


        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            this.Visible = true; 
            if (this.WindowState == FormWindowState.Minimized)
                this.WindowState = FormWindowState.Normal; // 최소화를 멈춘다
            this.Activate(); 
        }

  
        private void MainForm1_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true; // 종료 이벤트를 취소 시킨다
            this.Visible = false; // 폼을 표시하지 않는다;
        }

        //트레이의 종료메뉴 클릭시
        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {

            //트레이아이콘 없앰
            notifyIcon1.Visible = false;

            foreach (Process process in Process.GetProcesses()) { 
                if(process.ProcessName.ToUpper().StartsWith("AFA"))
                {
                    SetSystemCursor(LoadCursorFromFile("C:\\Users\\Administrator\\Desktop\\AFA_1.1_Server\\PPT_TEST\\aero_arrow.cur"), 32512);
                    process.Kill();

                }
            }
            Application.Exit();

            
        }

        
        //Bluetooth
        private void startScan()
        {
            //listBox1.DataSource = null;
            //listBox1.Items.Clear();
            items.Clear();
            Thread bluetoothScanThread = new Thread(new ThreadStart(scan));
            bluetoothScanThread.Start();
        }

        BluetoothDeviceInfo[] devices;

        private void scan()
        {
            updateUI("Starting Scan..");
            BluetoothClient client = new BluetoothClient();
            devices = client.DiscoverDevicesInRange();
            updateUI("Scan complete");
            updateUI(devices.Length.ToString() + " devices discovered");

            foreach (BluetoothDeviceInfo d in devices)
            {
                items.Add(d.DeviceName);
            }

            updateDeviceList();
        }

        private void connectAsServer()
        {
            Thread bluetoothServerThread = new Thread(new ThreadStart(ServerConnectThread));
            bluetoothServerThread.Start();
        }

        private void connectAsClient()
        {
            throw new NotImplementedException();
        }

        Guid mUUID = new Guid("00001101-0000-1000-8000-00805F9B34FB");
        bool serverStarted = false;

        public void ServerConnectThread()
        {
            serverStarted = true;
            updateUI("Server started, waiting for clients");
            blueListener = new BluetoothListener(mUUID);
            blueListener.Start();
            BluetoothClient conn = blueListener.AcceptBluetoothClient();
            updateUI("Client has connected");

            Stream mStream = conn.GetStream();

            while (true)
            {
                try
                {
                    //handle server connection
                    byte[] received = new byte[1024];
                    string received_msg;
                    mStream.Read(received, 0, received.Length);
                    received_msg = Encoding.ASCII.GetString(received);

                    //Bluetooth protocol parsing
                    parse_received(received_msg);
                }
                catch (IOException)
                {
                    updateUI("Client has disconnected!!!");
                }
            }
        }

        private void updateUI(string message)
        {
            Func<int> del = delegate()
            {
                tbOutput.AppendText(message + System.Environment.NewLine);
                return 0;
            };
            Invoke(del);
        }

        private void updateDeviceList()
        {
            Func<int> del = delegate()
            {
                //listBox1.DataSource = items;
                return 0;
            };
            Invoke(del);
        }

        BluetoothDeviceInfo deviceInfo;
        //private void listBox1_DoubleClick(object sender, EventArgs e)
        //{
        //    deviceInfo = devices.ElementAt(listBox1.SelectedIndex);
        //    updateUI(deviceInfo.DeviceName + " was selected, attempting connect");

        //    if (pairDevice())
        //    {
        //        updateUI("device paired..");
        //        updateUI("starting connect thread");
        //        Thread bluetoothClientThread = new Thread(new ThreadStart(ClientConnectThread));
        //        bluetoothClientThread.Start();
        //    }
        //    else
        //    {
        //        updateUI("Pair failed");
        //    }
        //}

        private void ClientConnectThread()
        {
            BluetoothClient client = new BluetoothClient();
            updateUI("attempting connect");
            client.BeginConnect(deviceInfo.DeviceAddress, mUUID, this.BluetoothClientConnectCallback, client);
        }

        void BluetoothClientConnectCallback(IAsyncResult result)
        {
            BluetoothClient client = (BluetoothClient)result.AsyncState;
            client.EndConnect(result);

            Stream stream = client.GetStream();
            stream.ReadTimeout = 1000;

            while (true)
            {
                while (!ready) ;

                stream.Write(message, 0, message.Length);
            }
        }

        string myPin = "1234";
        private bool pairDevice()
        {
            if (!deviceInfo.Authenticated)
            {
                if (!BluetoothSecurity.PairRequest(deviceInfo.DeviceAddress, myPin))
                {
                    return false;
                }
            }
            return true;
        }


        bool ready = false;
        byte[] message;

        //private void tbText_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    if (e.KeyChar == 13)
        //    {
        //        message = Encoding.ASCII.GetBytes(tbText.Text);
        //        ready = true;
        //        tbText.Clear();
        //    }
        //}



        /// <summary>
        /// Bluetooth Protocol 
        /// '1' - 좌하 
        /// '2' - 하
        /// '3' - 우하
        /// '4' - 좌
        /// '5' - 영점초기화
        /// '6' - 우
        /// '7' - 좌상
        /// '8' - 상
        /// '9' - 우상
        /// 'L' - 마우스 좌클릭
        /// 'R' - 마우스 우클릭
        /// 'P' - ppt Slide prev
        /// 'N' - ppt Slide next
        /// 'S' - ppt Slide Show, 커서변경(레이저포인터)
        /// 'C' - 커서변경(마우스포인터)
        /// 'D' - 클릭다운상태
        /// </summary>
        void parse_received(string msg)
        {
            //updateUI("protocol : " + msg);
            //updateUI("\r\n");

            int px = 15;
            int mx = -15;
            int py = 15;
            int my = -15;
  
            switch (msg[0]) { 
                case '1':
                    move_mouse(mx, py);
                    break;
                case '2':
                    move_mouse(0, py);
                    break;
                case '3':
                    move_mouse(px, py);
                    break;
                case '4':
                    move_mouse(mx, 0);
                    break;
                case '5':
                    x_prev = 900;
                    y_prev = 500;
                    SetCursorPos(900, 500);
                    break;
                case '6':
                    move_mouse(px, 0);
                    break;
                case '7':
                    move_mouse(mx, my);
                    break;
                case '8':
                    move_mouse(0, my);
                    break;
                case '9':
                    move_mouse(px, my);
                    break;
                case 'L':
                    mouse_Lclick();
                    break;
                case 'R':
                    mouse_Rclick();
                    break;
                case 'P':
                    slide_prev();
                    break;
                case 'N':
                    slide_next();
                    break;
                case 'D':
                    mouse_Dclick();
                    break;
                case 'S':
                    ppt_start();
                    break;
                case 'C':
                    SetSystemCursor(LoadCursorFromFile("C:\\Users\\Administrator\\Desktop\\AFA_1.1_Server\\PPT_TEST\\aero_arrow.cur"), 32512);
                    break;
                case 'b':
                    SetSystemCursor(LoadCursorFromFile("C:\\Users\\Administrator\\Desktop\\AFA_1.1_Server\\PPT_TEST\\ssbp1.cur"), 32512);
                    break;
                case 'B':
                    SetSystemCursor(LoadCursorFromFile("C:\\Users\\Administrator\\Desktop\\AFA_1.1_Server\\PPT_TEST\\ssbp3.cur"), 32512);
                    break;
                case 'y':
                    SetSystemCursor(LoadCursorFromFile("C:\\Users\\Administrator\\Desktop\\AFA_1.1_Server\\PPT_TEST\\syp1.cur"), 32512);
                    break;
                case'Y':
                    SetSystemCursor(LoadCursorFromFile("C:\\Users\\Administrator\\Desktop\\AFA_1.1_Server\\PPT_TEST\\syp3.cur"), 32512);
                    break;
                case 'r':
                    SetSystemCursor(LoadCursorFromFile("C:\\Users\\Administrator\\Desktop\\AFA_1.1_Server\\PPT_TEST\\srp1.cur"), 32512);
                    break;
                case 'E':
                    SetSystemCursor(LoadCursorFromFile("C:\\Users\\Administrator\\Desktop\\AFA_1.1_Server\\PPT_TEST\\srp2.cur"), 32512);
                    break;
            }
        }


        
        void move_mouse(int x, int y)
        {
            //updateUI("X : " + x);
            //updateUI("Y : " + y);
            //updateUI("\r\n");

            //포지션을 가져
            int xpos =0;
            int ypos =0;
            xpos = x_prev + x;
            ypos = y_prev + y;

            //마우스 이동
            SetCursorPos(xpos ,ypos);

            x_prev = xpos;
            y_prev = ypos;
        }

        void slide_prev()
        {
            slideIndex = slide.SlideIndex - 1;
            if (slideIndex >= 1)
            {
                try
                {
                    slide = slides[slideIndex];
                    slides[slideIndex].Select();
                }
                catch
                {
                    pptApplication.SlideShowWindows[1].View.Previous();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
            else
            {
                //MessageBox.Show("It is already Fist Page");

            }
        }

        void slide_next()
        {
            slideIndex = slide.SlideIndex + 1;
            if (slideIndex > slidescount)
            {
                //MessageBox.Show("It is already last page");
            }
            else
            {
                try
                {
                    slide = slides[slideIndex];
                    slides[slideIndex].Select();
                }
                catch
                {
                    pptApplication.SlideShowWindows[1].View.Next();
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }
        }

        void mouse_Lclick()
        {
            Point p = new Point(MousePosition.X, MousePosition.Y);
            mouse_event((int)MouseFlags.LEFTDOWN, (uint)p.X, (uint)p.Y, 0, 0);
            mouse_event((int)MouseFlags.LEFTUP, (uint)p.X, (uint)p.Y, 0, 0);
        }

        void mouse_Rclick()
        {
            Point p = new Point(MousePosition.X, MousePosition.Y);
            mouse_event((int)MouseFlags.RIGHTDOWN, (uint)p.X, (uint)p.Y, 0, 0);
            mouse_event((int)MouseFlags.RIGHTUP, (uint)p.X, (uint)p.Y, 0, 0);
        }

        void mouse_Dclick()
        {
            Point p = new Point(MousePosition.X, MousePosition.Y);
            mouse_event((int)MouseFlags.LEFTDOWN, (uint)p.X, (uint)p.Y, 0, 0);
        }

        void ppt_start()
        {
            try
            {
                // Get Running PowerPoint Application object 
                pptApplication = Marshal.GetActiveObject("PowerPoint.Application") as PPt.Application;

                // Get PowerPoint application successfully, then set control button enable 
                this.btnFirst.Enabled = true;
                this.btnNext.Enabled = true;
                this.btnPrev.Enabled = true;
                this.btnLast.Enabled = true;

                //마우스 커스 변경
                SetSystemCursor(LoadCursorFromFile("C:\\Users\\Administrator\\Desktop\\AFA_1.1_Server\\PPT_TEST\\laser.cur"), 32512);
                keybd_event((byte)Keys.F5, 0x45, 0, 0);
                keybd_event((byte)Keys.F5, 0x45, 0x02, 0);
            }
            catch
            {
                MessageBox.Show("Please Run PowerPoint Firstly", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
            }
            if (pptApplication != null)
            {
                // Get Presentation Object 
                presentation = pptApplication.ActivePresentation;
                // Get Slide collection object 
                slides = presentation.Slides;
                // Get Slide count 
                slidescount = slides.Count;
                // Get current selected slide  

                try
                {
                    // Get selected slide object in normal view 
                    slide = slides[pptApplication.ActiveWindow.Selection.SlideRange.SlideNumber];
                }
                catch
                {
                    // Get selected slide object in reading view 
                    slide = pptApplication.SlideShowWindows[1].View.Slide;
                }
            }


        }

        //마우스 커서 변경
        private void btnMouse_Click(object sender, EventArgs e)
        {
            SetSystemCursor(LoadCursorFromFile("C:\\Users\\Administrator\\Desktop\\AFA_1.1_Server\\PPT_TEST\\laser.cur"), 32512);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SetSystemCursor(LoadCursorFromFile("C:\\Users\\Administrator\\Desktop\\AFA_1.1_Server\\PPT_TEST\\aero_arrow.cur"), 32512);
        }
    }
}
