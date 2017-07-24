using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using System.IO;
using System.Threading;

namespace ModBusApp
{
    public partial class Form1 : Form
    {
        
        modbus mb = new modbus();
        string[] possiblePorts;
        int pollRate = 30000; //default of 30 seconds 
        int delay = 0;
        string saveLocation = "";
        string fileName = "log.csv";
        string fileName2 = "log2.csv";
        bool pollingState = false;
        bool loggingState = false;
        System.Threading.Timer timer;

        public Form1()
        {
            InitializeComponent();
           
            // Create the delegate that invokes methods for the timer.
            TimerCallback timerDelegate = new TimerCallback(timerElapsed);

            // Create a timer for polling:
            timer = new System.Threading.Timer(timerDelegate);
            
            //Create array of all available Serial Ports
            possiblePorts = SerialPort.GetPortNames();

            foreach (string port in possiblePorts)
            {
                serialPortComboBox.Items.Add(port);
            }//end for each

            //ComboBox Setup
            baudComboBox.Items.Add(9600);
            baudComboBox.Items.Add(14400);
            baudComboBox.Items.Add(19200);
            baudComboBox.Items.Add(38400);
            baudComboBox.Items.Add(57600);
            baudComboBox.Items.Add(115200);

            parityComboBox.Items.Add(Parity.Even);
            parityComboBox.Items.Add(Parity.Odd);
            parityComboBox.Items.Add(Parity.None);

            stopComboBox.Items.Add(StopBits.One);
            stopComboBox.Items.Add(StopBits.Two);

            dataComboBox.Items.Add(7);
            dataComboBox.Items.Add(8);


            valueComboBox1.Items.Add("HEX");
            valueComboBox1.Items.Add("DEC");
            valueComboBox1.Items.Add("BIN");
            valueComboBox1.SelectedItem = "HEX";

            valueComboBox2.Items.Add("HEX");
            valueComboBox2.Items.Add("DEC");
            valueComboBox2.Items.Add("BIN");
            valueComboBox2.SelectedItem = "HEX";

            valueComboBox3.Items.Add("HEX");
            valueComboBox3.Items.Add("DEC");
            valueComboBox3.Items.Add("BIN");
            valueComboBox3.SelectedItem = "HEX";

            valueComboBox4.Items.Add("HEX");
            valueComboBox4.Items.Add("DEC");
            valueComboBox4.Items.Add("BIN");
            valueComboBox4.SelectedItem = "HEX";

            valueComboBox5.Items.Add("HEX");
            valueComboBox5.Items.Add("DEC");
            valueComboBox5.Items.Add("BIN");
            valueComboBox5.SelectedItem = "HEX";

            valueComboBox6.Items.Add("HEX");
            valueComboBox6.Items.Add("DEC");
            valueComboBox6.Items.Add("BIN");
            valueComboBox6.SelectedItem = "HEX";

            valueComboBox7.Items.Add("HEX");
            valueComboBox7.Items.Add("DEC");
            valueComboBox7.Items.Add("BIN");
            valueComboBox7.SelectedItem = "HEX";

            valueComboBox8.Items.Add("HEX");
            valueComboBox8.Items.Add("DEC");
            valueComboBox8.Items.Add("BIN");
            valueComboBox8.SelectedItem = "HEX";

            valueComboBox9.Items.Add("HEX");
            valueComboBox9.Items.Add("DEC");
            valueComboBox9.Items.Add("BIN");
            valueComboBox9.SelectedItem = "HEX";

            valueComboBox10.Items.Add("HEX");
            valueComboBox10.Items.Add("DEC");
            valueComboBox10.Items.Add("BIN");
            valueComboBox10.SelectedItem = "HEX";

            

        }//end Form1()

        #region Polling
        private void timerElapsed(Object state)
        {

            if (InvokeRequired)
            {
                MethodInvoker method = new MethodInvoker(timerRead);
                Invoke(method);
                return;


            }//end if invoke required
            else
                timerRead();

        }//end timerElapsed()
      

        private void timerRead()
        {
            string name = "";
            //If "all registers" selected:
            if (allRadioButton.Checked)
            {
                for (int i = 3; i < 11; i++)
                {
                    Control con = this.Controls["groupBox1"];

                    name = "registerTextBox" + (i);
                    if (((TextBox)(this.Controls["groupBox1"].Controls[name])).Text != "")
                    {
                        ReadRegister input = new ReadRegister(((NumericUpDown)con.Controls["addressUpDown" + i]).Text, ((TextBox)con.Controls[name]).Text, (TextBox)con.Controls["responseTextBox" + i], "HEX", ((CheckBox)con.Controls["loggingCheckBox" + i]).Checked);
                        read(input);
                    }//end if

                }//end for

            }//end if all 

            //If "select registers" selected:
            else
            {
                for (int i = 3; i < 11; i++)
                {
                    Control con = this.Controls["groupBox1"];

                    name = "autoCheckBox" + (i-2);
                    if (((CheckBox)(this.Controls["groupBox1"].Controls[name])).Checked)
                    {
                        ReadRegister input = new ReadRegister(((NumericUpDown)con.Controls["addressUpDown" + i]).Text, ((TextBox)con.Controls["registerTextBox" + (i)]).Text, (TextBox)con.Controls["responseTextBox" + i], "HEX", ((CheckBox)con.Controls["loggingCheckBox" + i]).Checked);
                        read(input);
                    }//end if

                }//end for


            }//end selected

        }//end timerRead()
        #endregion

        #region Button Functions
        // Opens a file browser to select location for logs to be saved:
        private void browseButton_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                this.saveLocationTextBox.Text = folderBrowserDialog1.SelectedPath;
            }//end if result ok

        }//end browseButton_Click()

        // Changes the save location based on the textbox content:
        private void saveLocationTextBox_TextChanged(object sender, EventArgs e)
        {
            saveLocation = saveLocationTextBox.Text;
        }//end saveLocationTextBox_TextChanged()

        // Ensures valid input for poll rate:
        private void pollRateTextBox_TextChanged(object sender, EventArgs e)
        {
            //Determine if text is an appropriate number
            int num;
            if (int.TryParse(pollRateTextBox.Text, out num))
            {
                pollRate = num;
            }
            else
            {
                if (string.IsNullOrEmpty(pollRateTextBox.Text))
                    delay = 0;
                else
                    MessageBox.Show("Not a valid numerical selection", "Delay Rate Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }//end pollRateTextBox_TextChanged()

        //Ensures valid input for poll rate:
        private void delayTextBox_TextChanged(object sender, EventArgs e)
        {
            //Determine if text is an appropriate number
            int num;
            if (int.TryParse(delayTextBox.Text, out num))
            {
                delay = num;
            }
            else
            {
                if (string.IsNullOrEmpty(delayTextBox.Text))
                    delay = 0;
                else
                    MessageBox.Show("Not a valid numerical selection", "Delay Rate Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }//end delayTextBox_TextChanged()

        private void pollingEnableButton_Click(object sender, EventArgs e)
        {
            if (pollingState)
            {
                pollingState = false;
                MessageBox.Show("Automatic Polling Disabled");
                timer.Change(-1, -1);

            }
            else
            {
                if (!allRadioButton.Checked && !selectRadioButton.Checked)
                    selectRadioButton.Checked = true;
                pollingState = true;
                MessageBox.Show("Automatic Polling Enabled.");
                timer.Change(0, pollRate);
            }

        }//end pollingEnableButton_Click()

        private void enableLoggingButton_Click(object sender, EventArgs e)
        {
            if (loggingState)
            {
                loggingState = false;
                MessageBox.Show("Logging Disabled");
            }
            else
            {
                //If neither option was chosen, default to select registers:
                if (!loggingSelectRadioButton.Checked && !loggingAllRadioButton.Checked)
                    loggingSelectRadioButton.Checked = true;
                loggingState = true;
                MessageBox.Show("Logging Enabled.");
               
            }
        }//end enableLoggingButton_Click()

        // Refreshes the list of serialPorts available in the dropdown menu:
        private void serialPortComboBox_DropDown(object sender, EventArgs e)
        {
            serialPortComboBox.Items.Clear();
            possiblePorts = SerialPort.GetPortNames();
            foreach (string port in possiblePorts)
            {
                serialPortComboBox.Items.Add(port);
            }//end for each
        }//end serialPortComboBox_DropDown()

        #endregion

        #region Write
        // Calls modbus write function if all necessary parameters != null:
        private void writeFunction(object sender, EventArgs e)
        {
            WriteRegister input;
            string[] commands;  

            if (((Button)sender).Name == "writeButton1")
            {
                commands = commandTextBox1.Text.Split(' ');
                input = new WriteRegister(addressUpDown1.Text, registerTextBox1.Text, responseTextBox1, valueComboBox1.SelectedItem.ToString(), commandTextBox1.Text, loggingCheckBox1.Checked);
               
            }
            else
            {
                commands = commandTextBox2.Text.Split(' ');
                input = new WriteRegister(addressUpDown2.Text, registerTextBox2.Text, responseTextBox2, valueComboBox2.SelectedItem.ToString(), commandTextBox2.Text, loggingCheckBox1.Checked);
            }

            if (commands.Length == 1)
            {
                write(input);
            }
            else
            {
                for (int i = 0; i < commands.Length; i++)
                {
                    input.command = commands[i];
                    write(input);
                }//end for each data value         
            }

        }//end writeFunction()

        // Calls modbus write function if all necessary parameters != null:
        private void write(WriteRegister input)
        {
            ushort reg = 0;
          
            byte[] responses = new byte[8];

            if (input.register != "" && input.command != "" && input.address != null && input.value != null)
            {

                //commands = input.command.Split(' ');
                short[] value = new short[1];
                try
                {
                    value[0] = Convert.ToInt16(input.command, 16);
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }

                try
                {
                    if (input.value == "HEX")
                        reg = (ushort)Convert.ToUInt32(input.register, 16);
                    else if (input.value == "BIN")
                        reg = (ushort)Convert.ToUInt32(input.register, 2);
                    else
                        reg = (ushort)Convert.ToUInt32(input.register, 10);

                    mb.Open((string)serialPortComboBox.SelectedItem, (int)baudComboBox.SelectedItem, (int)dataComboBox.SelectedItem, (Parity)parityComboBox.SelectedItem, (StopBits)stopComboBox.SelectedItem);
                    mb.write(Convert.ToByte(input.address), reg, 1, value, ref responses, ref input);
                    string ack = "";

                    //Code for displaying all of response
                    /*
                    for (int i = 0; i < responses.Length; i++)
                    {
                        if (input.value == "HEX")
                            ack += responses[i].ToString("X") + " ";
                        else if (input.value == "DEC")
                            ack += responses[i].ToString() + " ";
                        else
                            ack += Convert.ToString(responses[i], 2) + " ";
                    }
                   */

                   //print out only the data value of response:
                    try
                    {

                        if (input.value == "HEX")
                            ack += responses[3].ToString("X") + " ";
                        else if (input.value == "DEC")
                            ack += responses[3].ToString() + " ";
                        else
                            ack += Convert.ToString(responses[3], 2) + " ";
                    }
                    catch
                    {
                        MessageBox.Show("Error with write response.");
                    }
                    
                    input.response.Text = ack;
                    string message = "";
                    for (int i = 0; i < input.message.Length; i++)
                        message += input.message[i].ToString("X") + " ";
                    if ( (input.isLogging && loggingState) || (loggingState && loggingAllRadioButton.Checked) )
                        log(DateTime.Now + "," + message + "," + input.response.Text);


                }
                catch (Exception err)
                {
                    MessageBox.Show("Error in write function: " + err.Message, "Write Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    log(DateTime.Now + "," + input.message + "," + input.response.Text + "," + err.Message);
                    return;
                }//end catch

            }//end if values != null

            else
                MessageBox.Show("Enter all fields before attempting a write.", "Write Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

        }//end writeFunction()
        #endregion

        #region Read
        // Assigns input and output value locations based on button and calls read:
        private void readFunction(object sender, EventArgs e)
        {
            ReadRegister input;
            

            if (((Button)sender).Name == "readButton10")
            {
                input = new ReadRegister(addressUpDown10.Text, registerTextBox10.Text, responseTextBox10, valueComboBox10.SelectedItem.ToString(), loggingCheckBox10.Checked);
                read(input);
              
            }
            else if (((Button)sender).Name == "readButton9")
            {
                input = new ReadRegister(addressUpDown9.Text, registerTextBox9.Text, responseTextBox9, valueComboBox9.SelectedItem.ToString(), loggingCheckBox9.Checked);
                read(input);
                
            }
            else if (((Button)sender).Name == "readButton8")
            {
                input = new ReadRegister(addressUpDown8.Text, registerTextBox8.Text, responseTextBox8, valueComboBox8.SelectedItem.ToString(), loggingCheckBox8.Checked);
                read(input);
               
            }
            else if (((Button)sender).Name == "readButton7")
            {
                input = new ReadRegister(addressUpDown7.Text, registerTextBox7.Text, responseTextBox7, valueComboBox7.SelectedItem.ToString(),loggingCheckBox7.Checked);
                read(input);
                
            }
            else if (((Button)sender).Name == "readButton6")
            {
                input = new ReadRegister(addressUpDown6.Text, registerTextBox6.Text, responseTextBox6, valueComboBox6.SelectedItem.ToString(), loggingCheckBox6.Checked);
                read(input);
                
            }
            else if (((Button)sender).Name == "readButton5")
            {
                input = new ReadRegister(addressUpDown10.Text, registerTextBox5.Text, responseTextBox5, valueComboBox5.SelectedItem.ToString(), loggingCheckBox5.Checked);
                read(input);
            }
            else if (((Button)sender).Name == "readButton4")
            {
                input = new ReadRegister(addressUpDown4.Text, registerTextBox4.Text, responseTextBox4, valueComboBox4.SelectedItem.ToString(), loggingCheckBox4.Checked);
                read(input);
            }
            else if (((Button)sender).Name == "readButton3")
            {
                input = new ReadRegister(addressUpDown3.Text, registerTextBox3.Text, responseTextBox3, valueComboBox3.SelectedItem.ToString(), loggingCheckBox3.Checked);
                read(input);
            }
            else if (((Button)sender).Name == "readButton2")
            {
                input = new ReadRegister(addressUpDown2.Text, registerTextBox2.Text, responseTextBox2, valueComboBox2.SelectedItem.ToString(), loggingCheckBox2.Checked);
                read(input);
            }
            else
            {
                input = new ReadRegister(addressUpDown1.Text, registerTextBox1.Text, responseTextBox1, valueComboBox1.SelectedItem.ToString() , loggingCheckBox1.Checked);
                read(input);
            }

        }//end readFunction()

        // Reads from the modbus (readSingle) given input: 
        private void read( ReadRegister input)
        {
            
            short[] values = new short[1];
            if (serialPortComboBox.SelectedItem != null && baudComboBox.SelectedItem != null && dataComboBox.SelectedItem != null && parityComboBox.SelectedItem != null && stopComboBox.SelectedItem != null && input.address != "" && input.register != "" && input.value != null)
            {
                string hexAddress = (Convert.ToInt32(input.address)).ToString("X");

                try
                {
                    mb.Open((string)serialPortComboBox.SelectedItem, (int)baudComboBox.SelectedItem, (int)dataComboBox.SelectedItem, (Parity)parityComboBox.SelectedItem, (StopBits)stopComboBox.SelectedItem);
                    mb.read(Convert.ToByte(hexAddress, 16), (ushort)Convert.ToInt32(input.register, 16), 1, ref values, ref input);
                    input.response.Text = values[0].ToString();
                    if (input.value == "HEX")
                        input.response.Text = values[0].ToString();
                    else if (input.value == "DEC")
                    {
                        input.response.Text = (Convert.ToInt32(values[0].ToString(), 16)).ToString();
                    }
                    else if (input.value == "BIN")
                    {
                        input.response.Text = Convert.ToString(Convert.ToInt32(values[0].ToString(), 16), 2);
                    }
                    
                }//end try
                catch (Exception err)
                {
                    MessageBox.Show("Error in read function: " + err.Message, "Read Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    log(DateTime.Now + "," + input.message + "," + input.response.Text + "," + err.Message);
                    return;
                }

            }//end if all parameters are not null
            else
            {
                MessageBox.Show("Enter all fields before attempting a read.", "Read Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            string message = "";
            if (input.message != null)
            {
                for (int i = 0; i < input.message.Length; i++)
                    message += input.message[i].ToString("X") + " ";
            }
            if (input.isLogging)
                log(DateTime.Now + "," + message + "," + input.response.Text + "," + input.status);
            mb.Close();
        }//end readFunction()

        // Calls the modbus multiple read function:
        private void multipleRead(object sender, EventArgs e)
        {
            ReadRegister input;
            string[] commands;
            string response = "";

            if (((Button)sender).Name == "readButton2")
            {
                commands = commandTextBox2.Text.Split(' ');
                input = new ReadRegister(addressUpDown2.Text, registerTextBox2.Text, responseTextBox2, valueComboBox2.SelectedItem.ToString(), loggingCheckBox2.Checked);
            }
            else
            {
                commands = commandTextBox1.Text.Split(' ');
                input = new ReadRegister(addressUpDown1.Text, registerTextBox1.Text, responseTextBox1, valueComboBox1.SelectedItem.ToString(), loggingCheckBox1.Checked);
            }
            if (commands[0] == "")
                commands[0] = "1";
            if (commands[0] == "1")
                read(input);
            else //If there are multiple read registers
            {
                for (int i = 0; i < Convert.ToInt16(commands[0]); i++)
                {
                    input.register = (Convert.ToInt16(input.register) + i).ToString();
                    read(input);
                    response += input.response.Text + " ";
                }//end for each data value     

                input.response.Text = response;
            }//end else
        }//end multipleRead()

        #endregion

        #region Logging
        private void log( string message )
        {
            if (saveLocation == "")
            {
                if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                {
                    this.saveLocationTextBox.Text = folderBrowserDialog1.SelectedPath;
                }//end if result ok
            }

            if (!File.Exists(saveLocation+"\\"+fileName))
            {
                // Create a file to write to.
                using (StreamWriter sw = File.AppendText(saveLocation + "\\" + fileName))
                {
                    sw.WriteLine("DateTime, Command, Response, Error");
                }//end using
            }//end if does not exist

            try
            {
                using (StreamWriter sw = File.AppendText(saveLocation + "\\" + fileName))
                {
                    sw.WriteLine(message);

                }//end using
            }
            catch (Exception ex)
            {
                MessageBox.Show("Logging Error: "+ex.Message, "Log Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            
        }//end log
        #endregion

    }//end Form

}//end NameSpace

