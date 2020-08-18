using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SZRFUI.Models
{
    public class EpsonRC90
    {
        public delegate void PrintEventHandler(object sender, string ModelMessageStr);
        public event PrintEventHandler EpsonStatusUpdate;
        public event EventHandler<bool> ConnectStateChanged;
        private bool mConnect = false;
        private bool _Connect
        {
            get { return mConnect; }
            set
            {
                if (mConnect != value)
                {
                    mConnect = value;
                    ConnectStateChanged?.Invoke(null, mConnect);
                }
            }
        }
        public TcpIpClient IOReceiveNet;
        public TcpIpClient CtrlNet;
        public string IP { get; set; }
        public int PORT { get; set; }
        public bool[] Rc90In;//从机械手读出
        public bool[] Rc90Out;//向机械手写入
        bool isLogined = false;
        public EpsonRC90(string ip,int port)
        {
            IOReceiveNet = new TcpIpClient();
            CtrlNet = new TcpIpClient();
            Rc90In = new bool[100];
            Rc90Out = new bool[100];
            IP = ip;
            PORT = port;
        }
        public async void checkIOReceiveNet()
        {
            while (true)
            {
                await Task.Delay(400);
                if (!IOReceiveNet.tcpConnected)
                {
                    await Task.Delay(1000);
                    if (!IOReceiveNet.tcpConnected)
                    {
                        bool r1 = await IOReceiveNet.Connect(IP, PORT);
                        if (r1)
                        {
                            _Connect = true;
                            // ModelPrint("机械手IOReceiveNet连接");

                        }
                        else
                            _Connect = false;
                    }
                }
                else
                { await Task.Delay(15000); }
            }
        }
        public async void IORevAnalysis()
        {
            while (true)
            {
                //await Task.Delay(100);
                try
                {
                    if (_Connect == true)
                    {
                        string s = await IOReceiveNet.ReceiveAsync();

                        string[] ss = s.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                        try
                        {
                            s = ss[0];

                        }
                        catch
                        {
                            s = "error";
                        }

                        if (s == "error")
                        {
                            IOReceiveNet.tcpConnected = false;
                            _Connect = false;
                            //ModelPrint("机械手IOReceiveNet断开");
                        }
                        else
                        {
                            string[] strs = s.Split(',');
                            if (strs[0] == "IOCMD" && strs[1].Length == 100)
                            {
                                for (int i = 0; i < 100; i++)
                                {
                                    Rc90In[i] = strs[1][i] == '1' ? true : false;
                                }
                                string RsedStr = "";
                                for (int i = 0; i < 100; i++)
                                {
                                    RsedStr += Rc90Out[i] ? "1" : "0";
                                }
                                await IOReceiveNet.SendAsync(RsedStr);
                                //ModelPrint("IOSend " + RsedStr);
                                //await Task.Delay(1);
                            }
                            //ModelPrint("IORev: " + s);
                        }
                    }
                    else
                    {
                        await Task.Delay(100);
                    }
                }
                catch { }
            }
        }
        public async void checkCtrlNet()
        {
            while (true)
            {

                if (!CtrlNet.tcpConnected)
                {
                    await Task.Delay(1000);
                    if (!CtrlNet.tcpConnected)
                    {
                        isLogined = false;
                        bool r1 = await CtrlNet.Connect(IP, 5000);
                        if (r1)
                        {
                            _Connect = true;
                        }
                        else
                            _Connect = false;
                    }
                }
                if (!isLogined && _Connect)
                {
                    await CtrlNet.SendAsync("$login,123");
                    string s = await CtrlNet.ReceiveAsync();
                    if (s.Contains("#login,0"))
                        isLogined = true;
                    await Task.Delay(400);
                }
                else
                {
                    await Task.Delay(3000);
                }
            }
        }
        public async void GetStatus()
        {
            string status = "";
            while (true)
            {
                await Task.Run(async ()=> {
                    if (isLogined == true)
                    {
                        try
                        {
                            status = await CtrlNet.SendAndReceive("$getstatus");
                            string[] statuss = status.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                            if (statuss[0] == "#getstatus")
                            {
                                if (statuss[1].Length == 11)
                                {

                                    EpsonStatusUpdate(null, statuss[1]);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            //Log.Default.Error("EpsonRC90.GetStatus", ex.Message);
                            await Task.Delay(1000);
                        }
                    }
                    //await Task.Delay(10);
                });
                
            }
        }
    }
}
