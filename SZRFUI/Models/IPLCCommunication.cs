using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SZRFUI.Models
{
    public interface IPLCCommunication
    {
        void SetM(string address, bool value);
        bool ReadM(string address);
        void SetMultiM(string address, bool[] value);
        bool[] ReadMultiM(string address, ushort length);
        /// <summary>
        /// 读单字
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        int ReadD(string address);
        /// <summary>
        /// 写单字
        /// </summary>
        /// <param name="address"></param>
        /// <param name="value"></param>
        void WriteD(string address, short value);
        /// <summary>
        /// 写多个单字
        /// </summary>
        /// <param name="address"></param>
        /// <param name="value"></param>
        void WriteMultD(string address, short[] value);
        /// <summary>
        /// 写双字
        /// </summary>
        /// <param name="address"></param>
        /// <param name="value"></param>
        void WriteW(string address, int value);
        /// <summary>
        /// 写多个双字
        /// </summary>
        /// <param name="address"></param>
        /// <param name="value"></param>
        void WriteMultW(string address, int[] value);
        /// <summary>
        /// 读单字
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        int ReadW(string address);
        /// <summary>
        /// 读多个双字
        /// </summary>
        /// <param name="address"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        int[] ReadMultiW(string address, ushort length);
    }
}
