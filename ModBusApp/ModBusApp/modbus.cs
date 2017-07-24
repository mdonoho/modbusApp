using System;
using System.Collections.Generic;
using System.Text;
using System.IO.Ports;
using System.Windows.Forms;
using System.Threading;
using ModBusApp;

namespace ModBusApp
{
    class modbus
    {
        private SerialPort sp = new SerialPort();
        public string modbusStatus;

        #region Constructor / Deconstructor
        public modbus()
        {
        }
        ~modbus()
        {
        }
        #endregion

        #region Open / Close Procedures
        public bool Open(string portName, int baudRate, int databits, Parity parity, StopBits stopBits)
        {
            //Ensure port isn't already opened:
            if (!sp.IsOpen)
            {
                //Assign desired settings to the serial port:
                sp.PortName = portName;
                sp.BaudRate = baudRate;
                sp.DataBits = databits;
                sp.Parity = parity;
                sp.StopBits = stopBits;
                //These timeouts are default and cannot be editted through the class at this point:
                sp.ReadTimeout = 5000;
                sp.WriteTimeout = 5000;

                //sp.ReceivedBytesThreshold = 1;
                //sp.RtsEnable = true;

                try
                {
                    sp.Open();
                }
                catch (Exception err)
                {
                    MessageBox.Show("Erroring Opening Serial Port", "Serial Port Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    modbusStatus = "Error opening " + portName + ": " + err.Message;
                    return false;
                }
                // MessageBox.Show("Serial Port Opened Successfully", "Serial Port Success", MessageBoxButtons.OK, MessageBoxIcon.None);
                modbusStatus = portName + " opened successfully";
                return true;
            }
            else
            {
                modbusStatus = portName + " already opened";
                return false;
            }
        }
        public bool Close()
        {
            //Ensure port is opened before attempting to close:
            if (sp.IsOpen)
            {
                try
                {
                    sp.Close();
                }
                catch (Exception err)
                {
                    modbusStatus = "Error closing " + sp.PortName + ": " + err.Message;
                    return false;
                }
                modbusStatus = sp.PortName + " closed successfully";
                return true;
            }
            else
            {
                modbusStatus = sp.PortName + " is not open";
                return false;
            }
        }
        #endregion

        #region CRC Computation
        private void GetCRC(byte[] message, ref byte[] CRC)
        {
            //Function expects a modbus message of any length as well as a 2 byte CRC array in which to 
            //return the CRC values:

            ushort CRCFull = 0xFFFF;
            byte CRCHigh = 0xFF, CRCLow = 0xFF;
            char CRCLSB;

            for (int i = 0; i < (message.Length) - 2; i++)
            {
                CRCFull = (ushort)(CRCFull ^ message[i]);

                for (int j = 0; j < 8; j++)
                {
                    CRCLSB = (char)(CRCFull & 0x0001);
                    CRCFull = (ushort)((CRCFull >> 1) & 0x7FFF);

                    if (CRCLSB == 1)
                        CRCFull = (ushort)(CRCFull ^ 0xA001);
                }
            }
            CRC[1] = CRCHigh = (byte)((CRCFull >> 8) & 0xFF);
            CRC[0] = CRCLow = (byte)(CRCFull & 0xFF);
          
        }
        #endregion

        #region Build Message
        private void BuildMessage(byte address, byte type, ushort start, ushort registers, ref byte[] message)
        {
            //Array to receive CRC bytes:
            byte[] CRC = new byte[2];

            message[0] = address;
            message[1] = type;
            message[2] = (byte)(start >> 8);
            message[3] = (byte)start;
            message[4] = (byte)(registers >> 8);
            message[5] = (byte)registers;

            
            GetCRC(message, ref CRC);
            message[message.Length - 2] = CRC[0];
            message[message.Length - 1] = CRC[1];
            
 
        }
        #endregion

        #region Check Response
        private bool CheckResponse(byte[] response)
        {
            //Perform a basic CRC check:
            byte[] CRC = new byte[2];
            GetCRC(response, ref CRC);
            if (CRC[0] == response[response.Length - 2] && CRC[1] == response[response.Length - 1])
                return true;
            else
                return false;
        }
        #endregion

        #region Get Response
        private void GetResponse(ref byte[] response)
        {
            //There is a bug in .Net 2.0 DataReceived Event that prevents people from using this
            //event as an interrupt to handle data (it doesn't fire all of the time).  Therefore
            //we have to use the ReadByte command for a fixed length as it's been shown to be reliable.
            for (int i = 0; i < response.Length; i++)
            {
                
                response[i] = (byte)(sp.ReadByte());
                //MessageBox.Show(response[i].ToString("X"));
            }

        }
        #endregion

        #region Write Register
        public bool write(byte address, ushort start, ushort registers, short[] values, ref byte[] response, ref WriteRegister input)
        {
            //Ensure port is open:
            if (sp.IsOpen)
            {
                //Clear in/out buffers:
                sp.DiscardOutBuffer();
                sp.DiscardInBuffer();
                //Message is 1 addr + 1 fcn + 2 start + 2 reg + 1 count + 2 * reg vals + 2 CRC
                byte[] message = new byte[8];
                //Function 16 response is fixed at 8 bytes
                //byte[] response = new byte[8];

                message[0] = address;
                message[1] = 0x06;
                message[2] = (byte)(start >> 8);
                message[3] = (byte)start;
                message[4] = (byte)(values[0] >> 8);
                message[5] = (byte)values[0];

                byte[] CRC = new byte[2];
                GetCRC(message, ref CRC);
                message[message.Length - 2] = CRC[0];
                message[message.Length - 1] = CRC[1];

                input.message = message;
                //MessageBox.Show(message[0].ToString("X") + " " + message[1].ToString("X") + " " + message[2].ToString("X") + " " + message[3].ToString("X") + " " + message[4].ToString("X") + " " + message[5].ToString("X") + " " + message[6].ToString("X") + " " + message[7].ToString("X"));


                //Send Modbus message to Serial Port:
                try
                {

                    sp.Write(message, 0, message.Length);
                    GetResponse(ref response);
                }
                catch (Exception err)
                {
                    MessageBox.Show("Write Error: " + err.Message, "Write Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    modbusStatus = "Error in write event: " + err.Message;
                    input.status = err.Message;
                    return false;
                }

                input.recieved = response;

                //Evaluate message:
                if (CheckResponse(response))
                {

                    MessageBox.Show("Write Successful");
                    modbusStatus = "Write successful";
                    return true;
                }//end if successful resposne
                else
                {
                    MessageBox.Show("CRC Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    modbusStatus = "CRC error";
                    input.status = "CRC ERROR";
                    return false;
                }//end else 

            }//end if SP Open
            else
            {
                MessageBox.Show("Serial Port Not Open", "Serial Port Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                modbusStatus = "Serial port not open";
                input.status = "SERIAL PORT NOT OPEN";
                return false;
            }
        }//end write()
        #endregion

        #region Read Register
        public bool read(byte address, ushort start, ushort registers, ref short[] values, ref ReadRegister input )
        {
            //Ensure port is open:
            if (sp.IsOpen)
            {
                //Clear in/out buffers:
                sp.DiscardOutBuffer();
                sp.DiscardInBuffer();
                //request is always 8 bytes:
                byte[] message = new byte[8];
                //response buffer:
                byte[] response = new byte[5 + 2 * registers];
                //Build outgoing modbus message:
                BuildMessage(address, 0x03, start, registers, ref message);

                input.message = message;

                //Send modbus message to Serial Port:
                try
                {
                    sp.Write(message, 0, message.Length);
                    GetResponse(ref response);
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message, "Read Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    modbusStatus = "Error in read event: " + err.Message;
                    input.status = err.Message;
                    return false;
                }

                input.recieved = response;

                //Evaluate message:
                if (CheckResponse(response))
                {
                    //Return requested register values:
                    for (int i = 0; i < (response.Length - 5) / 2; i++)
                    {
                        values[i] = response[2 * i + 3];
                        values[i] <<= 8;
                        values[i] += response[2 * i + 4];
                    }
                    MessageBox.Show("READ SUCCESSUL");
                    modbusStatus = "Read successful";

                    return true;
                }
                else
                {
                    MessageBox.Show("CRC Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    modbusStatus = "CRC error";
                    input.status = "CRC ERROR";
                    return false;
                }
            }
            else
            {
                MessageBox.Show("Serial Port Not Open", "Serial Port Error", MessageBoxButtons.OK, MessageBoxIcon.None);
                modbusStatus = "Serial port not open";
                input.status = "SERIAL PORT NOT OPEN";
                return false;
            }

        }//end read()
        #endregion

    }
}
