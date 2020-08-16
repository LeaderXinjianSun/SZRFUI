using BingLibrary.hjb.file;
using DXH.Net;
using Microsoft.Practices.Prism.Commands;
using Microsoft.Practices.Prism.ViewModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using SZRFUI.Models;

namespace SZRFUI.ViewModels
{
    class MainWindowViewModel : NotificationObject
    {
        #region 属性绑定
        private string windowTitle;

        public string WindowTitle
        {
            get { return windowTitle; }
            set
            {
                windowTitle = value;
                this.RaisePropertyChanged("WindowTitle");
            }
        }
        private bool statusPLC;

        public bool StatusPLC
        {
            get { return statusPLC; }
            set
            {
                statusPLC = value;
                this.RaisePropertyChanged("StatusPLC");
            }
        }
        private bool statusRobot;

        public bool StatusRobot
        {
            get { return statusRobot; }
            set
            {
                statusRobot = value;
                this.RaisePropertyChanged("StatusRobot");
            }
        }
        private string messageStr;

        public string MessageStr
        {
            get { return messageStr; }
            set
            {
                messageStr = value;
                this.RaisePropertyChanged("MessageStr");
            }
        }
        private long cycle;

        public long Cycle
        {
            get { return cycle; }
            set
            {
                cycle = value;
                this.RaisePropertyChanged("Cycle");
            }
        }
        private bool statusTester;

        public bool StatusTester
        {
            get { return statusTester; }
            set
            {
                statusTester = value;
                this.RaisePropertyChanged("StatusTester");
            }
        }

        #endregion
        #region 方法绑定
        public DelegateCommand AppLoadedEventCommand { get; set; }
        public DelegateCommand AppClosedEventCommand { get; set; }
        public DelegateCommand<object> MenuActionCommand { get; set; }
        public DelegateCommand FuncCommand { get; set; }
        #endregion
        #region 变量
        DXHTCPServer tcpServer;
        H3U h3u;
        EpsonRC90 epsonRC90;
        string iniParameterPath = System.Environment.CurrentDirectory + "\\Parameter.ini";
        bool[] M3000;
        #endregion
        #region 构造函数
        public MainWindowViewModel()
        {
            tcpServer = new DXHTCPServer();
            string plc_ip = Inifile.INIGetStringValue(iniParameterPath, "System", "PLCIP", "192.168.1.13");
            h3u = new H3U(plc_ip);
            string robot_ip = Inifile.INIGetStringValue(iniParameterPath, "System", "RobotIP", "192.168.1.5");
            int robot_port = int.Parse(Inifile.INIGetStringValue(iniParameterPath, "System", "RobotPORT", "2000"));
            epsonRC90 = new EpsonRC90(robot_ip, robot_port);
            h3u.ModbusTCP_Client.ModbusStateChanged += ModbusTCP_Client_ModbusStateChanged;
            epsonRC90.ConnectStateChanged += EpsonRC90_ConnectStateChanged;


            tcpServer.LocalIPAddress = Inifile.INIGetStringValue(iniParameterPath, "System", "ServerIP", "127.0.0.1");
            tcpServer.LocalIPPort = int.Parse(Inifile.INIGetStringValue(iniParameterPath, "System", "ServerPORT", "11001"));
            tcpServer.SocketListChanged += TcpServer_SocketListChanged;
            tcpServer.ConnectStateChanged += TcpServer_ConnectStateChanged;
            tcpServer.Received += TcpServer_Received;


            AppLoadedEventCommand = new DelegateCommand(new Action(this.AppLoadedEventCommandExecute));
            AppClosedEventCommand = new DelegateCommand(new Action(this.AppClosedEventCommandExecute));
            MenuActionCommand = new DelegateCommand<object>(new Action<object>(this.MenuActionCommandExecute));
            FuncCommand = new DelegateCommand(new Action(this.FuncCommandExecute));
            if (System.Environment.CurrentDirectory != @"C:\Debug")
            {
                System.Windows.MessageBox.Show("软件安装目录必须为C:\\Debug", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                System.Windows.Application.Current.Shutdown();
            }
            else
            {
                System.Diagnostics.Process[] myProcesses = System.Diagnostics.Process.GetProcessesByName("SZRFUI");//获取指定的进程名   
                if (myProcesses.Length > 1) //如果可以获取到知道的进程名则说明已经启动
                {
                    System.Windows.MessageBox.Show("不允许重复打开软件", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                    System.Windows.Application.Current.Shutdown();
                }
            }

        }

        private void TcpServer_Received(object sender, string e)
        {
            string restr = e.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries)[0];
            AddMessage("接收:" + restr);
            try
            {
                string[] strs = restr.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                switch (strs[0])
                {
                    case "Heart":
                        Socket socket = tcpServer.SocketList.FirstOrDefault();
                        AddMessage(@"Heart,1");
                        string str = tcpServer.TCPSend(socket, @"Heart,1" + "\r\n", false);
                        if (str != "")
                        {
                            AddMessage(str);
                        }
                        break;
                    case "Result":
                        string[] strs1 = strs[1].Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                        if (strs1.Length > 1)
                        {
                            switch (strs1[1])
                            {
                                case "1"://B
                                    if (strs1[0] == "0")
                                    {
                                        h3u.SetM("M3103", true);
                                    }
                                    else
                                    {
                                        h3u.SetM("M3105", true);
                                    }
                                    break;
                                case "2"://A                                  
                                    if (strs1[0] == "0")
                                    {
                                        h3u.SetM("M3102", true);
                                    }
                                    else
                                    {
                                        h3u.SetM("M3104", true);
                                    }
                                    break;
                                default:
                                    break;
                            }
                        }
                        else//样本测试结果
                        {
                            if (strs[1] == "3")
                            {
                                h3u.SetM("M3117", true);
                            }
                            else
                            {
                                h3u.SetM("M3118", true);
                            }
                        }
                       
                        break;
                    case "Sample"://开始样本测试
                        h3u.SetM("M3115", true);
                        break;
                    case "Feed"://允许放料
                        h3u.SetM("M3116", true);
                        break;
                    case "Reset"://重启完成
                        switch (strs[1])
                        {
                            case "1":
                                h3u.SetM("M3110", true);
                                break;
                            case "2":
                                h3u.SetM("M3111", true);
                                break;
                            default:
                                break;
                        }
                        break;
                    case "Alarm":
                        string[] strs4 = strs[1].Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                        switch (strs4[1])
                        {
                            case "1"://B
                                switch (strs4[0])
                                {
                                    case "1"://真空报警
                                        AddMessage("B真空报警");
                                        h3u.SetM("M3107", true);
                                        break;
                                    case "2"://连续3次NG报警
                                        AddMessage("B连续3次NG报警");
                                        break;
                                    case "3"://气缸下到位NG
                                        AddMessage("B气缸下到位NG");
                                        break;
                                    case "4"://气缸上到位NG
                                        AddMessage("B气缸上到位NG");
                                        break;
                                    case "5"://光纤叠料
                                        AddMessage("B光纤叠料");
                                        h3u.SetM("M3109", true);
                                        break;
                                    case "6"://上位料掉料
                                        AddMessage("B上位料掉料");
                                        break;
                                    default:
                                        break;
                                }
                                break;
                            case "2"://A                                
                                switch (strs4[0])
                                {
                                    case "1"://真空报警
                                        AddMessage("A真空报警");
                                        h3u.SetM("M3106", true);
                                        break;
                                    case "2"://连续3次NG报警
                                        AddMessage("A连续3次NG报警");
                                        break;
                                    case "3"://气缸下到位NG
                                        AddMessage("A气缸下到位NG");
                                        break;
                                    case "4"://气缸上到位NG
                                        AddMessage("A气缸上到位NG");
                                        break;
                                    case "5"://光纤叠料
                                        AddMessage("A光纤叠料");
                                        h3u.SetM("M3108", true);
                                        break;
                                    case "6"://上位料掉料
                                        AddMessage("A上位料掉料");
                                        break;
                                    default:
                                        break;
                                }
                                break;
                            default:
                                break;
                        }
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                AddMessage(ex.Message);
            }
           
        }

        private void TcpServer_ConnectStateChanged(object sender, string e)
        {
            AddMessage("TCP Server " + e);
        }

        private void TcpServer_SocketListChanged(object sender, bool e)
        {
            AddMessage($"测试机:{((IPEndPoint)(((Socket)sender).RemoteEndPoint)).Address}:{((IPEndPoint)(((Socket)sender).RemoteEndPoint)).Port} {(e ? "连接" : "断开")}");
            StatusTester = e;
        }

        private void EpsonRC90_ConnectStateChanged(object sender, bool e)
        {
            StatusRobot = e;
        }

        private void ModbusTCP_Client_ModbusStateChanged(object sender, bool e)
        {
            StatusPLC = e;
        }
        #endregion
        #region 方法绑定函数
        private void AppLoadedEventCommandExecute()
        {
            Init();
        }
        private void AppClosedEventCommandExecute()
        {
            try
            {
                tcpServer.Close();
            }
            catch { }
        }
        private void MenuActionCommandExecute(object p)
        {

        }
        private void FuncCommandExecute()
        {

        }
        #endregion
        #region 功能函数
        private void Init()
        {
            WindowTitle = "SZRFUI20200814";
            MessageStr = "";
            StatusPLC = true;
            StatusRobot = true;
            StatusTester = true;
            h3u.ModbusTCP_Client.StartConnect();
            epsonRC90.checkIOReceiveNet();
            epsonRC90.IORevAnalysis();
            Task.Run(()=> { PLCRun(); });
            tcpServer.StartTCPListen();
            AddMessage("软件加载完成");
        }
        void PLCRun()
        {
            bool M3171 = false;
            Stopwatch sw = new Stopwatch();
            while (true)
            {
                sw.Restart();
                try
                {
                    #region 互刷
                    if (StatusPLC && StatusRobot)
                    {
                        M3000 = h3u.ReadMultiM("M3000", 50);
                        for (int i = 0; i < M3000.Length; i++)
                        {
                            epsonRC90.Rc90Out[i] = M3000[i];
                        }
                        h3u.SetMultiM("M3050", epsonRC90.Rc90In);
                        System.Threading.Thread.Sleep(20);
                    }
                    else
                    {
                        System.Threading.Thread.Sleep(1000);
                    }
                    #endregion
                    #region 与测试机交互                    
                    if (StatusPLC)
                    {
                        bool[] M3150 = h3u.ReadMultiM("M3150", 24);
                        if (M3150[2])
                        {
                            Socket socket = tcpServer.SocketList.FirstOrDefault();
                            AddMessage(@"Unload,0:1");
                            string str = tcpServer.TCPSend(socket, @"Unload,0:1" + "\r\n", false);
                            if (str != "")
                            {
                                AddMessage(str);
                            }
                            AddMessage(@"Unload,0:2");
                            str = tcpServer.TCPSend(socket, @"Unload,0:2" + "\r\n", false);
                            if (str != "")
                            {
                                AddMessage(str);
                            }
                            h3u.SetM("M3152", false);
                        }
                        
                        if (M3150[3])//复位
                        {
                            Socket socket = tcpServer.SocketList.FirstOrDefault();
                            AddMessage(@"Reboot");
                            string str = tcpServer.TCPSend(socket, @"Reboot" + "\r\n", false);
                            if (str != "")
                            {
                                AddMessage(str);
                            }
                            h3u.SetM("M3153", false);
                        }
                        
                        if (M3150[4])//重启
                        {
                            Socket socket = tcpServer.SocketList.FirstOrDefault();
                            AddMessage(@"Reset,0");
                            string str = tcpServer.TCPSend(socket, @"Reset,0" + "\r\n", false);
                            if (str != "")
                            {
                                AddMessage(str);
                            }
                            h3u.SetM("M3154", false);
                        }
                        
                        if (M3150[20])//放完成
                        {
                            Socket socket = tcpServer.SocketList.FirstOrDefault();
                            AddMessage(@"Feed,1:1");
                            string str = tcpServer.TCPSend(socket, @"Feed,1:1" + "\r\n", false);
                            if (str != "")
                            {
                                AddMessage(str);
                            }
                            AddMessage(@"Feed,1:2");
                            str = tcpServer.TCPSend(socket, @"Feed,1:2" + "\r\n", false);
                            if (str != "")
                            {
                                AddMessage(str);
                            }
                            h3u.SetM("M3170", false);
                        }
                        
                        bool m3171 = M3150[21];//开始放入样本产品
                        if (M3171 != m3171)
                        {
                            M3171 = m3171;
                            if (m3171)
                            {
                                Socket socket = tcpServer.SocketList.FirstOrDefault();
                                AddMessage(@"Sample,1");
                                string str = tcpServer.TCPSend(socket, @"Sample,1" + "\r\n", false);
                                if (str != "")
                                {
                                    AddMessage(str);
                                }
                            }
                        }
                        
                        if (M3150[22])//样本上料结束
                        {
                            Socket socket = tcpServer.SocketList.FirstOrDefault();
                            AddMessage(@"Sample,2");
                            string str = tcpServer.TCPSend(socket, @"Sample,2" + "\r\n", false);
                            if (str != "")
                            {
                                AddMessage(str);
                            }
                            h3u.SetM("M3172", false);
                        }
                        
                    }
                    #endregion
                }
                catch { System.Threading.Thread.Sleep(1000); }

                Cycle = sw.ElapsedMilliseconds;
            }
        }
        private void AddMessage(string str)
        {
            string[] s = MessageStr.Split('\n');
            if (s.Length > 1000)
            {
                MessageStr = "";
            }
            if (MessageStr != "")
            {
                MessageStr += "\n";
            }
            MessageStr += System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + " " + str;
            RunLog(str);
        }
        void RunLog(string str)
        {
            try
            {
                string tempSaveFilee5 = System.AppDomain.CurrentDomain.BaseDirectory + @"RunLog";
                DateTime dtim = DateTime.Now;
                string DateNow = dtim.ToString("yyyy/MM/dd");
                string TimeNow = dtim.ToString("HH:mm:ss");

                if (!Directory.Exists(tempSaveFilee5))
                {
                    Directory.CreateDirectory(tempSaveFilee5);  //创建目录 
                }

                if (File.Exists(tempSaveFilee5 + "\\" + DateNow.Replace("/", "") + ".txt"))
                {
                    //第一种方法：
                    FileStream fs = new FileStream(tempSaveFilee5 + "\\" + DateNow.Replace("/", "") + ".txt", FileMode.Append);
                    StreamWriter sw = new StreamWriter(fs);
                    sw.WriteLine("TTIME：" + TimeNow + " 执行事件：" + str);
                    sw.Dispose();
                    fs.Dispose();
                    sw.Close();
                    fs.Close();
                }
                else
                {
                    //不存在就新建一个文本文件,并写入一些内容 
                    StreamWriter sw;
                    sw = File.CreateText(tempSaveFilee5 + "\\" + DateNow.Replace("/", "") + ".txt");
                    sw.WriteLine("TTIME：" + TimeNow + " 执行事件：" + str);
                    sw.Dispose();
                    sw.Close();
                }
            }
            catch { }
        }
        #endregion
    }
}
