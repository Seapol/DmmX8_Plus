using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using Ivi.Visa.Interop;
using System.Diagnostics;
using Microsoft.CSharp.RuntimeBinder;
using System.Windows.Forms;

namespace CypressSemiconductor.ChinaManufacturingTest
{

    
    
    public class MultiMeter
    {
        private static FormattedIO488 ioMultiMeterCntrl;

        public static string devReadyStatus = "";

        public struct current
        {
            public double max;
            public double min;
            public double average;
        }



        public MultiMeter(string ConnectString)
        {

            devReadyStatus = InitializeU3606A(ConnectString);



        }

        ~MultiMeter()
        {
            ioMultiMeterCntrl = null;
        }

        
        #region U3606A Methods
        private bool U3606A_Connect(string connectString)
        {
            try
            {
                // connectString="USB0::2391::16664::MY48470014::0::INSTR";
                // Create the resource manager and open a session with the instrument specified on U2722AtxtAddr
                ResourceManager grm1 = new ResourceManager();
                ioMultiMeterCntrl.IO = (IMessage)grm1.Open(connectString, AccessMode.NO_LOCK, 2000, "");


                ioMultiMeterCntrl.IO.Timeout = 7000;
                // Only return true if previous calls have also returned true
                return true;

            }
            catch
            {
                ioMultiMeterCntrl.IO = null;
                return false;
            }
        }


        private string InitializeU3606A(string ConnectString)
        {
            bool devReady = false;
            string idnTemp;

            ioMultiMeterCntrl = new FormattedIO488();   // Create the formatted I/O object

            if (ConnectString=="")
            {
                //MessageBox.Show("Error Due to invalid instrument alias: " + ConnectString);
                return "Error Due to invalid instrument alias: null";
            }

            devReady = U3606A_Connect(ConnectString);
            if (devReady == false)
            {
                //throw new Exception("Connect to U3606A failed on " + ConnectString);
                //MessageBox.Show("Connect to U3606A failed on " + ConnectString);
                return "Connect to U3606A failed on " + ConnectString;
            }
            else
            {
                ioMultiMeterCntrl.WriteString("*RST", true);              // Reset
                ioMultiMeterCntrl.WriteString("*CLS", true);              // Clear
                ioMultiMeterCntrl.WriteString("*IDN?", true);             // Read ID string
                idnTemp = ioMultiMeterCntrl.ReadString();

                Trace.WriteLine("Connect to U3606A suceess! U3606A IDN: " + idnTemp.ToString());
                //MessageBox.Show("Connect to U3606A suceess! U3606A IDN: " + idnTemp.ToString());

                ioMultiMeterCntrl.WriteString("SYST:BEEP:STAT OFF", true);

                //ioMultiMeterCntrl.WriteString("CONF:CURR 0.01,0.000001", true); // 10mA range with 1uA resolution 
                //ioMultiMeterCntrl.WriteString("CONF:CURR 0.01,0.0000001", true);  // 10mA range with 0.1uA resolution

                //ioMultiMeterCntrl.WriteString("CONF:CURR 0.1,0.00001", true); // 100mA range with 10uA resolution 
                //ioMultiMeterCntrl.WriteString("CONF:CURR 0.1,0.000001", true);  // 100mA range with 1uA resolution

                ioMultiMeterCntrl.WriteString("CONF:CURR 1,0.0001", true);  // 1A range with 100uA resolution
                //ioMultiMeterCntrl.WriteString("CONF:CURR 1,0.00001", true);  // 1A range with 10uA resolution;
                /*
                 * By my test, "CONF:CURR 1, 0.00001" can't set resolution to 10uA. 
                 * But use the following command can't reset resolution to 10uA
                */
                //ioMultiMeterCntrl.WriteString("CURR:RES 1E-05", true);

                //ioMultiMeterCntrl.WriteString("CONF:CURR 3,0.001", true);  // 3A range with 1mA resolution
                //ioMultiMeterCntrl.WriteString("CONF:CURR 3,0.0001", true);  // 3A range with 100uA resolution; By my test it is not supported

                return "Connect to U3606A Successfully!!! U3606A IDN: " + idnTemp.ToString();
            }
        }

        public current MeasureChannelCurrent()
        {
            lock (ioMultiMeterCntrl)
            {
                current ch_current=new current();

                ioMultiMeterCntrl.WriteString("CALC:STAT ON", true);             //turn on the calculation function
                ioMultiMeterCntrl.WriteString("CALC:FUNC AVER", true);      //calculation function is Average


                System.Threading.Thread.Sleep(500); //wait for 500ms

                ioMultiMeterCntrl.WriteString("CALC:AVER:AVER?", true);
                ch_current.average = (double)ioMultiMeterCntrl.ReadNumber(IEEEASCIIType.ASCIIType_Any, true);

                ioMultiMeterCntrl.WriteString("CALC:AVER:MAX?", true);
                ch_current.max = (double)ioMultiMeterCntrl.ReadNumber(IEEEASCIIType.ASCIIType_Any, true);

                ioMultiMeterCntrl.WriteString("CALC:AVER:MIN?", true);
                ch_current.min = (double)ioMultiMeterCntrl.ReadNumber(IEEEASCIIType.ASCIIType_Any, true);


                ioMultiMeterCntrl.WriteString("CALC:STAT OFF", true);            //turn off the calculation function


                //ioMultiMeterCntrl.WriteString("FETC?", true);
                //ch_current = (double)ioMultiMeterCntrl.ReadNumber(IEEEASCIIType.ASCIIType_Any, true);

                return ch_current;
            }
        }

       
        #endregion

    }//end of class

}//end of namespace
