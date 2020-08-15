using DXH.Modbus;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SZRFUI.Models
{
    public class H3U : IPLCCommunication
    {
        public DXHModbusTCP ModbusTCP_Client;
        public H3U(string ip)
        {
            ModbusTCP_Client = new DXHModbusTCP();
            ModbusTCP_Client.RemoteIPAddress = ip;
            ModbusTCP_Client.RemoteIPPort = 502;
        }
        public int ReadD(string address)
        {
            try
            {
                if (address.Substring(0, 1) == "D")
                {
                    int _add = int.Parse(address.Substring(1, address.Length - 1));
                    int[] result = ModbusTCP_Client.ModbusRead(1, 3, _add);
                    return result[0];
                }
                else
                {
                    return -1;
                }
            }
            catch
            {
                return -1;
            }
        }

        public bool ReadM(string address)
        {
            try
            {
                if (address.Substring(0, 1) == "M")
                {
                    int _add = int.Parse(address.Substring(1, address.Length - 1));
                    int[] result = ModbusTCP_Client.ModbusRead(1, 1, _add);
                    return result[0] == 1;
                }
                else
                {
                    return false;
                }
            }
            catch
            {
                return false;
            }
        }

        public bool[] ReadMultiM(string address, ushort length)
        {
            try
            {
                if (address.Substring(0, 1) == "M")
                {
                    int _add = int.Parse(address.Substring(1, address.Length - 1));
                    int[] result = ModbusTCP_Client.ModbusRead(1, 1, _add,length);
                    bool[] values = new bool[length];
                    for (int i = 0; i < length; i++)
                    {
                        values[i] = result[i] == 1;
                    }
                    return values;
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                return null;
            }
        }

        public int[] ReadMultiW(string address, ushort length)
        {
            try
            {
                if (address.Substring(0, 1) == "D")
                {
                    int _add = int.Parse(address.Substring(1, address.Length - 1));
                    int[] result = ModbusTCP_Client.ModbusRead(1, 3, _add, length * 2);
                    int[] values = new int[length];
                    for (int i = 0; i < length; i++)
                    {
                        string coilvalue = result[i * 2 + 1].ToString("X4") + result[i * 2].ToString("X4");
                        values[i] = Convert.ToInt32(coilvalue, 16);
                    }
                    return values;
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                return null;
            }
        }

        public int ReadW(string address)
        {
            try
            {
                if (address.Substring(0, 1) == "D")
                {
                    int _add = int.Parse(address.Substring(1, address.Length - 1));
                    int[] result = ModbusTCP_Client.ModbusRead(1, 3, _add, 2);
                    string coilvalue = result[1].ToString("X4") + result[0].ToString("X4");
                    return Convert.ToInt32(coilvalue, 16);
                }
                else
                {
                    return -1;
                }
            }
            catch
            {
                return -1;
            }
        }

        public void SetM(string address, bool value)
        {
            try
            {
                if (address.Substring(0, 1) == "M")
                {
                    int _add = int.Parse(address.Substring(1, address.Length - 1));
                    int[] b = new int[1];
                    b[0] = value ? 1 : 0;
                    ModbusTCP_Client.ModbusWrite(1, 15, _add, b);//写线圈
                }     
            }
            catch
            {
                
            }
        }

        public void SetMultiM(string address, bool[] value)
        {
            try
            {
                if (address.Substring(0, 1) == "M")
                {
                    int _add = int.Parse(address.Substring(1, address.Length - 1));
                    int[] b = new int[value.Length];
                    for (int i = 0; i < value.Length; i++)
                    {
                        b[i] = value[i] ? 1 : 0;
                    }                    
                    ModbusTCP_Client.ModbusWrite(1, 15, _add, b);//写线圈
                }
            }
            catch
            {

            }
        }

        public void WriteD(string address, short value)
        {
            try
            {
                if (address.Substring(0, 1) == "D")
                {
                    int _add = int.Parse(address.Substring(1, address.Length - 1));
                    int[] b = new int[1];
                    b[0] = value;
                    ModbusTCP_Client.ModbusWrite(1, 16, _add, b);
                }
            }
            catch
            {

            }
        }

        public void WriteMultD(string address, short[] value)
        {
            try
            {
                if (address.Substring(0, 1) == "D")
                {
                    int _add = int.Parse(address.Substring(1, address.Length - 1));
                    int[] b = new int[value.Length];
                    for (int i = 0; i < value.Length; i++)
                    {
                        b[i] = value[i];
                    }
                    ModbusTCP_Client.ModbusWrite(1, 16, _add, b);//写线圈
                }
            }
            catch
            {

            }
        }

        public void WriteMultW(string address, int[] value)
        {
            try
            {
                if (address.Substring(0, 1) == "D")
                {
                    int _add = int.Parse(address.Substring(1, address.Length - 1));
                    int[] b = new int[2 * value.Length];
                    for (int i = 0; i < value.Length; i++)
                    {
                        b[2 * i] = Convert.ToInt32(value[i].ToString("X8").Substring(4, 4), 16);
                        b[2 * i + 1] = Convert.ToInt32(value[i].ToString("X8").Substring(0, 4), 16);
                    }
                    ModbusTCP_Client.ModbusWrite(1, 16, _add, b);
                }
            }
            catch
            {

            }
        }

        public void WriteW(string address, int value)
        {
            try
            {
                if (address.Substring(0, 1) == "D")
                {
                    int _add = int.Parse(address.Substring(1, address.Length - 1));
                    int[] b = new int[2];
                    b[0] = Convert.ToInt32(value.ToString("X8").Substring(4, 4), 16);
                    b[1] = Convert.ToInt32(value.ToString("X8").Substring(0, 4), 16);
                    ModbusTCP_Client.ModbusWrite(1, 16, _add, b);
                }
            }
            catch
            {

            }
        }
    }
}
