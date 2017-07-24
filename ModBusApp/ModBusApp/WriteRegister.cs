using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ModBusApp
{
    class WriteRegister
    {
        public string address { get; set; }
        public string register { get; set; }
        public TextBox response { get; set; }
        public string value { get; set; }
        public string command { get; set; }
        public byte[] message { get; set; }
        public bool isLogging { get; set; }
        public string status { get; set; }
        public byte[] recieved { get; set; }

        #region Constructor / Deconstructor
        public WriteRegister()
        {
        }
        public WriteRegister(string add, string reg, TextBox resp, string val, string comm, bool log)
        {
            address = add;
            register = reg;
            response = resp;
            value = val;
            command = comm;
            isLogging = log;
        }
        ~WriteRegister()
        {
        }
        #endregion
    }
}
