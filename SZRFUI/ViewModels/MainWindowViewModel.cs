using BingLibrary.hjb.file;
using DXH.Net;
using Microsoft.Practices.Prism.Commands;
using Microsoft.Practices.Prism.ViewModel;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
        private bool epsonStatusAuto;

        public bool EpsonStatusAuto
        {
            get { return epsonStatusAuto; }
            set
            {
                epsonStatusAuto = value;
                this.RaisePropertyChanged("EpsonStatusAuto");
            }
        }
        private bool epsonStatusWarning;

        public bool EpsonStatusWarning
        {
            get { return epsonStatusWarning; }
            set
            {
                epsonStatusWarning = value;
                this.RaisePropertyChanged("EpsonStatusWarning");
            }
        }
        private bool epsonStatusSError;

        public bool EpsonStatusSError
        {
            get { return epsonStatusSError; }
            set
            {
                epsonStatusSError = value;
                this.RaisePropertyChanged("EpsonStatusSError");
            }
        }
        private bool epsonStatusSafeGuard;

        public bool EpsonStatusSafeGuard
        {
            get { return epsonStatusSafeGuard; }
            set
            {
                epsonStatusSafeGuard = value;
                this.RaisePropertyChanged("EpsonStatusSafeGuard");
            }
        }
        private bool epsonStatusEStop;

        public bool EpsonStatusEStop
        {
            get { return epsonStatusEStop; }
            set
            {
                epsonStatusEStop = value;
                this.RaisePropertyChanged("EpsonStatusEStop");
            }
        }
        private bool epsonStatusError;

        public bool EpsonStatusError
        {
            get { return epsonStatusError; }
            set
            {
                epsonStatusError = value;
                this.RaisePropertyChanged("EpsonStatusError");
            }
        }
        private bool epsonStatusPaused;

        public bool EpsonStatusPaused
        {
            get { return epsonStatusPaused; }
            set
            {
                epsonStatusPaused = value;
                this.RaisePropertyChanged("EpsonStatusPaused");
            }
        }
        private bool epsonStatusRunning;

        public bool EpsonStatusRunning
        {
            get { return epsonStatusRunning; }
            set
            {
                epsonStatusRunning = value;
                this.RaisePropertyChanged("EpsonStatusRunning");
            }
        }
        private bool epsonStatusReady;

        public bool EpsonStatusReady
        {
            get { return epsonStatusReady; }
            set
            {
                epsonStatusReady = value;
                this.RaisePropertyChanged("EpsonStatusReady");
            }
        }
        private ObservableCollection<AlarmReportFormViewModel> alarmReportForm;

        public ObservableCollection<AlarmReportFormViewModel> AlarmReportForm
        {
            get { return alarmReportForm; }
            set
            {
                alarmReportForm = value;
                this.RaisePropertyChanged("AlarmReportForm");
            }
        }
        private string homePageVisibility;

        public string HomePageVisibility
        {
            get { return homePageVisibility; }
            set
            {
                homePageVisibility = value;
                this.RaisePropertyChanged("HomePageVisibility");
            }
        }
        private string alarmReportFormPageVisibility;

        public string AlarmReportFormPageVisibility
        {
            get { return alarmReportFormPageVisibility; }
            set
            {
                alarmReportFormPageVisibility = value;
                this.RaisePropertyChanged("AlarmReportFormPageVisibility");
            }
        }

        #endregion
        #region 方法绑定
        public DelegateCommand AppLoadedEventCommand { get; set; }
        public DelegateCommand AppClosedEventCommand { get; set; }
        public DelegateCommand<object> MenuActionCommand { get; set; }
        public DelegateCommand FuncCommand { get; set; }
        public DelegateCommand<object> EpsonOperateCommand { get; set; }
        public DelegateCommand AlarmReportFromExportCommand { get; set; }
        #endregion
        #region 变量
        DXHTCPServer tcpServer;
        Fx5u fx5u;
        EpsonRC90 epsonRC90;
        string iniParameterPath = System.Environment.CurrentDirectory + "\\Parameter.ini";
        bool[] M2300;
        Metro metro = new Metro();
        List<AlarmData> AlarmList = new List<AlarmData>();
        string LastBanci;
        #endregion
        #region 构造函数
        public MainWindowViewModel()
        {
            tcpServer = new DXHTCPServer();
            string plc_ip = Inifile.INIGetStringValue(iniParameterPath, "System", "PLCIP", "192.168.1.13");
            int plc_port = int.Parse(Inifile.INIGetStringValue(iniParameterPath, "System", "PLCPORT", "3900"));
            fx5u = new Fx5u(plc_ip, plc_port);
            string robot_ip = Inifile.INIGetStringValue(iniParameterPath, "System", "RobotIP", "192.168.1.5");
            int robot_port = int.Parse(Inifile.INIGetStringValue(iniParameterPath, "System", "RobotPORT", "2000"));
            epsonRC90 = new EpsonRC90(robot_ip, robot_port);
            fx5u.ConnectStateChanged += ModbusTCP_Client_ModbusStateChanged;
            epsonRC90.ConnectStateChanged += EpsonRC90_ConnectStateChanged;
            epsonRC90.EpsonStatusUpdate += EpsonRC90_EpsonStatusUpdate; ;

            tcpServer.LocalIPAddress = Inifile.INIGetStringValue(iniParameterPath, "System", "ServerIP", "127.0.0.1");
            tcpServer.LocalIPPort = int.Parse(Inifile.INIGetStringValue(iniParameterPath, "System", "ServerPORT", "11001"));
            tcpServer.SocketListChanged += TcpServer_SocketListChanged;
            tcpServer.ConnectStateChanged += TcpServer_ConnectStateChanged;
            tcpServer.Received += TcpServer_Received;


            AppLoadedEventCommand = new DelegateCommand(new Action(this.AppLoadedEventCommandExecute));
            AppClosedEventCommand = new DelegateCommand(new Action(this.AppClosedEventCommandExecute));
            MenuActionCommand = new DelegateCommand<object>(new Action<object>(this.MenuActionCommandExecute));
            EpsonOperateCommand = new DelegateCommand<object>(new Action<object>(this.EpsonOperateCommandExecute));
            AlarmReportFromExportCommand = new DelegateCommand(new Action(this.AlarmReportFromExportCommandExecute));
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

        private void EpsonRC90_EpsonStatusUpdate(object sender, string str)
        {
            try
            {
                EpsonStatusAuto = str[2] == '1';
                EpsonStatusWarning = str[3] == '1';
                EpsonStatusSError = str[4] == '1';
                EpsonStatusSafeGuard = str[5] == '1';
                EpsonStatusEStop = str[6] == '1';
                EpsonStatusError = str[7] == '1';
                EpsonStatusPaused = str[8] == '1';
                EpsonStatusRunning = str[9] == '1';
                EpsonStatusReady = str[10] == '1';

                bool[] robotStatus = new bool[] { EpsonStatusAuto, EpsonStatusWarning, EpsonStatusSError, EpsonStatusSafeGuard,
                        EpsonStatusEStop,EpsonStatusError,EpsonStatusPaused,EpsonStatusRunning,EpsonStatusReady};
                if (StatusPLC)
                {
                    fx5u.SetMultiM("M405", robotStatus);
                }
            }
            catch { }
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
                                        fx5u.SetM("M2505", true);
                                    }
                                    else
                                    {
                                        fx5u.SetM("M2506", true);
                                    }
                                    break;
                                case "2"://A                                  
                                    if (strs1[0] == "0")
                                    {
                                        fx5u.SetM("M2503", true);
                                    }
                                    else
                                    {
                                        fx5u.SetM("M2504", true);
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
                                fx5u.SetM("M2520", true);
                            }
                            else
                            {
                                fx5u.SetM("M2521", true);
                            }
                        }
                       
                        break;
                    case "Sample"://开始样本测试
                        fx5u.SetM("M2510", true);
                        break;
                    case "Feed"://允许放料
                        fx5u.SetM("M2501", true);
                        break;
                    case "Reset"://重启完成
                        switch (strs[1])
                        {
                            case "1":
                                fx5u.SetM("M2525", true);
                                break;
                            case "2":
                                fx5u.SetM("M2524", true);
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
                                        fx5u.SetM("M2536", true);
                                        break;
                                    case "2"://连续3次NG报警
                                        AddMessage("B连续3次NG报警");
                                        fx5u.SetM("M2537", true);
                                        break;
                                    case "3"://气缸下到位NG
                                        AddMessage("B气缸下到位NG");
                                        fx5u.SetM("M2538", true);
                                        break;
                                    case "4"://气缸上到位NG
                                        AddMessage("B气缸上到位NG");
                                        fx5u.SetM("M2539", true);
                                        break;
                                    case "5"://光纤叠料
                                        AddMessage("B光纤叠料");
                                        fx5u.SetM("M2540", true);
                                        break;
                                    case "6"://上位料掉料
                                        AddMessage("B上位料掉料");
                                        fx5u.SetM("M2541", true);
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
                                        fx5u.SetM("M2526", true);
                                        break;
                                    case "2"://连续3次NG报警
                                        AddMessage("A连续3次NG报警");
                                        fx5u.SetM("M2527", true);
                                        break;
                                    case "3"://气缸下到位NG
                                        AddMessage("A气缸下到位NG");
                                        fx5u.SetM("M2528", true);
                                        break;
                                    case "4"://气缸上到位NG
                                        AddMessage("A气缸上到位NG");
                                        fx5u.SetM("M2529", true);
                                        break;
                                    case "5"://光纤叠料
                                        AddMessage("A光纤叠料");
                                        fx5u.SetM("M2530", true);
                                        break;
                                    case "6"://上位料掉料
                                        AddMessage("A上位料掉料");
                                        fx5u.SetM("M2531", true);
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
            switch (p.ToString())
            {
                case "0":
                    HomePageVisibility = "Visible";
                    AlarmReportFormPageVisibility = "Collapsed";
                    break;
                case "1":
                    break;
                case "2":
                    HomePageVisibility = "Collapsed";
                    AlarmReportFormPageVisibility = "Visible";
                    break;
                default:
                    break;
            }
        }
        private void FuncCommandExecute()
        {

        }
        private async void EpsonOperateCommandExecute(object p)
        {
            switch (p.ToString())
            {
                case "0":
                    if (StatusRobot && EpsonStatusReady && !EpsonStatusEStop)
                    {
                        await epsonRC90.CtrlNet.SendAsync("$start,0");
                    }
                    break;
                case "1":
                    if (StatusRobot)
                    {
                        await epsonRC90.CtrlNet.SendAsync("$pause");
                    }
                    break;
                case "2":
                    if (StatusRobot)
                    {
                        await epsonRC90.CtrlNet.SendAsync("$continue");
                    }
                    break;
                case "3":
                    metro.ChangeAccent("Dark.Red");
                    bool r = await metro.ShowConfirm("确认", "你确定重启机械手吗?");
                    metro.ChangeAccent("Light.Blue");
                    if (r && StatusRobot)
                    {
                        await epsonRC90.CtrlNet.SendAsync("$stop");
                        await Task.Delay(300);
                        await epsonRC90.CtrlNet.SendAsync("$SetMotorOff,1");
                        await Task.Delay(400);
                        await epsonRC90.CtrlNet.SendAsync("$reset");
                    }

                    break;
                default:
                    break;
            }
        }
        private void AlarmReportFromExportCommandExecute()
        {
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            //dlg.FileName = "AlarmReport"; // Default file name
            //dlg.DefaultExt = ".xlsx"; // Default file extension
            dlg.Filter = "Text Files(*.xlsx)|*.xlsx|All(*.*)|*"; // Filter files by extension

            // Show save file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                // Save document
                WriteAlarmtoExcel(dlg.FileName);
            }
        }
        #endregion
        #region 功能函数
        private void Init()
        {
            WindowTitle = "SZRFUI20200816:N104";
            MessageStr = "";
            HomePageVisibility = "Visible";
            AlarmReportFormPageVisibility = "Collapsed";
            StatusPLC = true;
            StatusRobot = true;
            StatusTester = true;
            LastBanci = Inifile.INIGetStringValue(iniParameterPath, "Summary", "LastBanci", "null");
            epsonRC90.checkIOReceiveNet();
            epsonRC90.IORevAnalysis();
            epsonRC90.checkCtrlNet();
            epsonRC90.GetStatus();
            Task.Run(()=> { PLCRun(); });
            tcpServer.StartTCPListen();
            UIRun();
            
            #region 加载报警记录
            try
            {
                using (StreamReader reader = new StreamReader(System.IO.Path.Combine(System.Environment.CurrentDirectory, "AlarmReportForm.json")))
                {
                    string json = reader.ReadToEnd();
                    AlarmReportForm = JsonConvert.DeserializeObject<ObservableCollection<AlarmReportFormViewModel>>(json);
                }
            }
            catch (Exception ex)
            {
                AlarmReportForm = new ObservableCollection<AlarmReportFormViewModel>();
                WriteToJson(AlarmReportForm, System.IO.Path.Combine(System.Environment.CurrentDirectory, "AlarmReportForm.json"));
                AddMessage(ex.Message);
            }
            #endregion
            #region 报警文档
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                string alarmExcelPath = Path.Combine(System.Environment.CurrentDirectory, "RF上料机报警.xlsx");
                if (File.Exists(alarmExcelPath))
                {

                    FileInfo existingFile = new FileInfo(alarmExcelPath);
                    using (ExcelPackage package = new ExcelPackage(existingFile))
                    {
                        // get the first worksheet in the workbook
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
                        {
                            AlarmData ad = new AlarmData();
                            ad.Code = worksheet.Cells["A" + i.ToString()].Value == null ? "Null" : worksheet.Cells["A" + i.ToString()].Value.ToString();
                            ad.Content = worksheet.Cells["B" + i.ToString()].Value == null ? "Null" : worksheet.Cells["B" + i.ToString()].Value.ToString();
                            ad.Type = worksheet.Cells["C" + i.ToString()].Value == null ? "Null" : worksheet.Cells["C" + i.ToString()].Value.ToString();
                            ad.Start = DateTime.Now;
                            ad.End = DateTime.Now;
                            ad.State = false;
                            AlarmList.Add(ad);
                        }
                        AddMessage("读取到" + worksheet.Dimension.End.Row.ToString() + "条报警");
                    }
                }
                else
                {
                    AddMessage("VPP报警.xlsx 文件不存在");
                }
            }
            catch (Exception ex)
            {
                AddMessage(ex.Message);
            }
            #endregion


            AddMessage("软件加载完成");
        }
        void PLCRun()
        {
            bool M2607 = false;
            Stopwatch sw = new Stopwatch();
            while (true)
            {
                sw.Restart();
                try
                {
                    #region 互刷
                    if (StatusPLC && StatusRobot)
                    {
                        M2300 = fx5u.ReadMultiM("M2300", 100);
                        for (int i = 0; i < M2300.Length; i++)
                        {
                            epsonRC90.Rc90Out[i] = M2300[i];
                        }
                        fx5u.SetMultiM("M2200", epsonRC90.Rc90In);
                        System.Threading.Thread.Sleep(100);
                    }
                    else
                    {
                        System.Threading.Thread.Sleep(1000);
                    }
                    #endregion
                    #region 与测试机交互                    
                    if (StatusPLC)
                    {
                        bool[] M2600 = fx5u.ReadMultiM("M2600", 32);
                        if (M2600[3])
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
                            fx5u.SetM("M2603", false);
                        }
                        
                        if (M2600[16])//复位
                        {
                            Socket socket = tcpServer.SocketList.FirstOrDefault();
                            AddMessage(@"Reboot");
                            string str = tcpServer.TCPSend(socket, @"Reboot" + "\r\n", false);
                            if (str != "")
                            {
                                AddMessage(str);
                            }
                            fx5u.SetM("M2616", false);
                        }
                        
                        if (M2600[24])//重启
                        {
                            Socket socket = tcpServer.SocketList.FirstOrDefault();
                            AddMessage(@"Reset,0");
                            string str = tcpServer.TCPSend(socket, @"Reset,0" + "\r\n", false);
                            if (str != "")
                            {
                                AddMessage(str);
                            }
                            fx5u.SetM("M2624", false);
                        }
                        
                        if (M2600[2])//放完成
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
                            fx5u.SetM("M2602", false);
                        }
                        
                        bool m2607 = M2600[7];//开始放入样本产品
                        if (M2607 != m2607)
                        {
                            M2607 = m2607;
                            if (M2607)
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
                        
                        if (M2600[9])//样本上料结束
                        {
                            Socket socket = tcpServer.SocketList.FirstOrDefault();
                            AddMessage(@"Sample,2");
                            string str = tcpServer.TCPSend(socket, @"Sample,2" + "\r\n", false);
                            if (str != "")
                            {
                                AddMessage(str);
                            }
                            fx5u.SetM("M2609", false);
                        }
                        
                    }
                    #endregion
                }
                catch { System.Threading.Thread.Sleep(1000); }

                Cycle = sw.ElapsedMilliseconds;
            }
        }
        private async void UIRun()
        {
            string CurrentAlarm = "";
            if (!Directory.Exists("D:\\报警记录"))
            {
                Directory.CreateDirectory("D:\\报警记录");
            }
            while (true)
            {
                #region PLC后台操作机械手

                try
                {
                    bool[] M420 = await Task.Run<bool[]>(() => { return fx5u.ReadMultiM("M420", 4); });
                    if (M420 != null && StatusRobot)
                    {
                        if (M420[0])
                        {
                            await epsonRC90.CtrlNet.SendAsync("$start,0");
                            //AddMessage("$start,0");
                        }
                        if (M420[1])
                        {
                            await epsonRC90.CtrlNet.SendAsync("$pause");
                            //AddMessage("$pause");
                        }
                        if (M420[2])
                        {
                            await epsonRC90.CtrlNet.SendAsync("$continue");
                            //AddMessage("$continue");
                        }
                        if (M420[3])
                        {
                            await epsonRC90.CtrlNet.SendAsync("$stop");
                            //AddMessage("$stop");
                            await Task.Delay(300);
                            await epsonRC90.CtrlNet.SendAsync("$SetMotorOff,1");
                            //AddMessage("$SetMotorOff,1");
                            await Task.Delay(400);
                            await epsonRC90.CtrlNet.SendAsync("$reset");
                            //AddMessage("$reset");
                        }
                    }


                }
                catch { }
                #endregion
                #region 报警记录
                try
                {
                    //读报警
                    bool[] M300 = fx5u.ReadMultiM("M300", (ushort)AlarmList.Count);
                    if (M300 != null && StatusPLC)
                    {
                        for (int i = 0; i < AlarmList.Count; i++)
                        {
                            if (M300[i] != AlarmList[i].State && AlarmList[i].Content != "Null")
                            {
                                AlarmList[i].State = M300[i];
                                if (AlarmList[i].State)
                                {
                                    AlarmList[i].Start = DateTime.Now;
                                    AlarmList[i].End = DateTime.Now;
                                    AddMessage(AlarmList[i].Code + AlarmList[i].Content + "发生");
                                    var nowAlarm = AlarmReportForm.FirstOrDefault(s => s.Code == AlarmList[i].Code);
                                    if (nowAlarm == null)
                                    {
                                        AlarmReportFormViewModel newAlarm = new AlarmReportFormViewModel()
                                        {
                                            Code = AlarmList[i].Code,
                                            Content = AlarmList[i].Content,
                                            Count = 1,
                                            TimeSpan = AlarmList[i].End - AlarmList[i].Start
                                        };
                                        AlarmReportForm.Add(newAlarm);
                                        WriteToJson(AlarmReportForm, System.IO.Path.Combine(System.Environment.CurrentDirectory, "AlarmReportForm.json"));
                                    }
                                    else
                                    {
                                        if (CurrentAlarm != AlarmList[i].Content)
                                        {
                                            nowAlarm.Count++;
                                            WriteToJson(AlarmReportForm, System.IO.Path.Combine(System.Environment.CurrentDirectory, "AlarmReportForm.json"));
                                            CurrentAlarm = AlarmList[i].Content;
                                        }
                                    }                                    
                                }
                                else
                                {
                                    AlarmList[i].End = DateTime.Now;
                                    AddMessage(AlarmList[i].Code + AlarmList[i].Content + "解除");
                                    var nowAlarm = AlarmReportForm.FirstOrDefault(s => s.Code == AlarmList[i].Code);
                                    if (nowAlarm != null)
                                    {
                                        nowAlarm.TimeSpan += AlarmList[i].End - AlarmList[i].Start;
                                        WriteToJson(AlarmReportForm, System.IO.Path.Combine(System.Environment.CurrentDirectory, "AlarmReportForm.json"));
                                    }
                                }
                            }
                        }

                    }
                }
                catch (Exception ex)
                {
                    AddMessage(ex.Message);
                }
                #endregion
                #region 换班
                if (LastBanci != GetBanci())
                {
                    try
                    {
                        WriteAlarmtoExcel(Path.Combine("D:\\报警记录", "RF报警统计" + LastBanci + ".xlsx"));
                        AlarmReportForm.Clear();
                        WriteToJson(AlarmReportForm, System.IO.Path.Combine(System.Environment.CurrentDirectory, "AlarmReportForm.json"));
                        LastBanci = GetBanci();
                        Inifile.INIWriteValue(iniParameterPath, "Summary", "LastBanci", LastBanci);
                        AddMessage(LastBanci + " 换班数据清零");
                    }
                    catch (Exception ex)
                    {
                        AddMessage(ex.Message);
                    }
                }
                #endregion
                await Task.Delay(200);
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
        private void WriteAlarmtoExcel(string filepath)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage())
                {
                    var ws = package.Workbook.Worksheets.Add("MySheet");
                    ws.Cells[1, 1].Value = "ID";
                    ws.Cells[1, 2].Value = "报警内容";
                    ws.Cells[1, 3].Value = "报警次数";
                    ws.Cells[1, 4].Value = "报警时长(分钟)";
                    ws.Cells[1, 5].Value = DateTime.Now.ToString();
                    for (int i = 0; i < AlarmReportForm.Count; i++)
                    {
                        ws.Cells[i + 2, 1].Value = AlarmReportForm[i].Code;
                        ws.Cells[i + 2, 2].Value = AlarmReportForm[i].Content;
                        ws.Cells[i + 2, 3].Value = AlarmReportForm[i].Count;
                        ws.Cells[i + 2, 4].Value = Math.Round(AlarmReportForm[i].TimeSpan.TotalMinutes, 1);
                    }
                    package.SaveAs(new FileInfo(filepath));
                }

            }
            catch (Exception ex)
            {
                AddMessage(ex.Message);
            }

        }
        private void WriteToJson(object p, string path)
        {
            try
            {
                using (FileStream fs = File.Open(path, FileMode.Create))
                using (StreamWriter sw = new StreamWriter(fs))
                using (JsonWriter jw = new JsonTextWriter(sw))
                {
                    jw.Formatting = Formatting.Indented;
                    JsonSerializer serializer = new JsonSerializer();
                    serializer.Serialize(jw, p);
                }
            }
            catch (Exception ex)
            {
                AddMessage(ex.Message);
            }
        }
        private string GetBanci()
        {
            string rs = "";
            if (DateTime.Now.Hour >= 8 && DateTime.Now.Hour < 20)
            {
                rs += DateTime.Now.ToString("yyyyMMdd") + "_D";
            }
            else
            {
                if (DateTime.Now.Hour >= 0 && DateTime.Now.Hour < 8)
                {
                    rs += DateTime.Now.AddDays(-1).ToString("yyyyMMdd") + "_N";
                }
                else
                {
                    rs += DateTime.Now.ToString("yyyyMMdd") + "_N";
                }
            }
            return rs;
        }
        #endregion
    }
}
