using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using CypressSemiconductor.ChinaManufacturingTest;


namespace DmmX8_Plus
{
    public partial class Form1 : Form
    {

        MultiMeter DMM;
        Agilent SW_A;
        bool devReady = false;
        bool channelClosed = false;

        public int SNLength = DmmX8_Plus.Properties.Settings.Default.SNLength;
        

        public Form1()
        {
            InitializeComponent();

            channelBox1.Text = "CH1";
            channelBox2.Text = "CH2";
            channelBox3.Text = "CH3";
            channelBox4.Text = "CH4";
            channelBox5.Text = "CH5";
            channelBox6.Text = "CH6";
            channelBox7.Text = "CH7";
            channelBox8.Text = "CH8";



            DMM = Initialize_U3606A(DmmX8_Plus.Properties.Settings.Default.DMM_Alias);

            if (MultiMeter.devReadyStatus.Contains("Connect to U3606A Successfully!!!"))
            {
                dmm_status.Text = "Ready";
                dmm_status.ForeColor = Color.DarkGreen;
                devReady = true;
                checkBox1.Checked = true;
            }
            else
            {
                MessageBox.Show("Connect to U3606A failed on " + DmmX8_Plus.Properties.Settings.Default.DMM_Alias, "Instrument Initializaion Error");
                checkBox1.Checked = false;
            }

            SW_A = Initialize_U2751A(DmmX8_Plus.Properties.Settings.Default.Switch_Alias);

            if (Agilent.devReadyStatus == true)
            {
                sw_status.Text = "Ready";
                sw_status.ForeColor = Color.DarkGreen;
                devReady = true;
                checkBox2.Checked = true;

            }
            else
            {
                MessageBox.Show("Connect to U2751A failed on " + DmmX8_Plus.Properties.Settings.Default.Switch_Alias, "Instrument Initializaion Error");
                checkBox2.Checked = false;

            }



            serialnumberbox1.Focus();

        }

        private void DoWork1()
        {
            this.Invoke(new MethodInvoker(() => toolStripStatusLabel1.Text = "Running"));
            Thread.Sleep(100);
            this.Invoke(new MethodInvoker(() => toolStripProgressBar1.Value = 30));
            this.Invoke(new MethodInvoker(() => RunTest(channelBox1.Text, toolStripProgressBar1, measuredlabel1, TestResultLabel1, toolStripStatusLabel1)));
            this.Invoke(new MethodInvoker(() => toolStripStatusLabel1.Text = "Done"));
            this.Invoke(new MethodInvoker(() => serialnumberbox1.Text = ""));


        }

        private void DoWork2()
        {
            this.Invoke(new MethodInvoker(() => toolStripStatusLabel2.Text = "Running"));
            Thread.Sleep(100);
            this.Invoke(new MethodInvoker(() => toolStripProgressBar2.Value = 30));
            this.Invoke(new MethodInvoker(() => RunTest(channelBox2.Text, toolStripProgressBar2, measuredlabel2, TestResultLabel2, toolStripStatusLabel2)));
            this.Invoke(new MethodInvoker(() => toolStripStatusLabel2.Text = "Done"));
            this.Invoke(new MethodInvoker(() => serialnumberbox2.Text = ""));


        }

        private void DoWork3()
        {
            this.Invoke(new MethodInvoker(() => toolStripStatusLabel3.Text = "Running"));
            Thread.Sleep(100);
            this.Invoke(new MethodInvoker(() => toolStripProgressBar3.Value = 30));
            this.Invoke(new MethodInvoker(() => RunTest(channelBox3.Text, toolStripProgressBar3, measuredlabel3, TestResultLabel3, toolStripStatusLabel3)));
            this.Invoke(new MethodInvoker(() => toolStripStatusLabel3.Text = "Done"));
            this.Invoke(new MethodInvoker(() => serialnumberbox3.Text = ""));


        }

        private void DoWork4()
        {
            this.Invoke(new MethodInvoker(() => toolStripStatusLabel4.Text = "Running"));
            Thread.Sleep(100);
            this.Invoke(new MethodInvoker(() => toolStripProgressBar4.Value = 30));
            this.Invoke(new MethodInvoker(() => RunTest(channelBox4.Text, toolStripProgressBar4, measuredlabel4, TestResultLabel4, toolStripStatusLabel4)));
            this.Invoke(new MethodInvoker(() => toolStripStatusLabel4.Text = "Done"));
            this.Invoke(new MethodInvoker(() => serialnumberbox4.Text = ""));

        }

        private void DoWork5()
        {
            this.Invoke(new MethodInvoker(() => toolStripStatusLabel5.Text = "Running"));
            Thread.Sleep(100);
            this.Invoke(new MethodInvoker(() => toolStripProgressBar5.Value = 30));
            this.Invoke(new MethodInvoker(() => RunTest(channelBox5.Text, toolStripProgressBar5, measuredlabel5, TestResultLabel5, toolStripStatusLabel5)));
            this.Invoke(new MethodInvoker(() => toolStripStatusLabel5.Text = "Done"));
            this.Invoke(new MethodInvoker(() => serialnumberbox5.Text = ""));

        }

        private void DoWork6()
        {
            this.Invoke(new MethodInvoker(() => toolStripStatusLabel6.Text = "Running"));
            Thread.Sleep(100);
            this.Invoke(new MethodInvoker(() => toolStripProgressBar6.Value = 30));
            this.Invoke(new MethodInvoker(() => RunTest(channelBox6.Text, toolStripProgressBar6, measuredlabel6, TestResultLabel6, toolStripStatusLabel6)));
            this.Invoke(new MethodInvoker(() => toolStripStatusLabel6.Text = "Done"));
            this.Invoke(new MethodInvoker(() => serialnumberbox6.Text = ""));

        }

        private void DoWork7()
        {
            this.Invoke(new MethodInvoker(() => toolStripStatusLabel7.Text = "Running"));
            Thread.Sleep(100);
            this.Invoke(new MethodInvoker(() => toolStripProgressBar7.Value = 30));
            this.Invoke(new MethodInvoker(() => RunTest(channelBox7.Text, toolStripProgressBar7, measuredlabel7, TestResultLabel7, toolStripStatusLabel7)));
            this.Invoke(new MethodInvoker(() => toolStripStatusLabel7.Text = "Done"));
            this.Invoke(new MethodInvoker(() => serialnumberbox7.Text = ""));

        }

        private void DoWork8()
        {
            this.Invoke(new MethodInvoker(() => toolStripStatusLabel8.Text = "Running"));
            Thread.Sleep(100);
            this.Invoke(new MethodInvoker(() => toolStripProgressBar8.Value = 30));
            this.Invoke(new MethodInvoker(() => RunTest(channelBox8.Text, toolStripProgressBar8, measuredlabel8, TestResultLabel8, toolStripStatusLabel8)));
            this.Invoke(new MethodInvoker(() => toolStripStatusLabel8.Text = "Done"));
            this.Invoke(new MethodInvoker(() => serialnumberbox8.Text = ""));

        }

        private void run1_Click(object sender, EventArgs e)
        {
            Thread t = new Thread(DoWork1);
            t.IsBackground = true;
            t.Start();

        }

        private void run2_Click(object sender, EventArgs e)
        {
            Thread t = new Thread(DoWork2);
            t.IsBackground = true;
            t.Start();
        }

        private void run3_Click(object sender, EventArgs e)
        {
            Thread t = new Thread(DoWork3);
            t.IsBackground = true;
            t.Start();
        }

        private void run4_Click(object sender, EventArgs e)
        {
            Thread t = new Thread(DoWork4);
            t.IsBackground = true;
            t.Start();
        }

        private void run5_Click(object sender, EventArgs e)
        {
            Thread t = new Thread(DoWork5);
            t.IsBackground = true;
            t.Start();
        }

        private void run6_Click(object sender, EventArgs e)
        {
            Thread t = new Thread(DoWork6);
            t.IsBackground = true;
            t.Start();
        }

        private void run7_Click(object sender, EventArgs e)
        {
            Thread t = new Thread(DoWork7);
            t.IsBackground = true;
            t.Start();
        }

        private void run8_Click(object sender, EventArgs e)
        {
            Thread t = new Thread(DoWork8);
            t.IsBackground = true;
            t.Start();
        }

        private void allrun_Click(object sender, EventArgs e)
        {
            if (channelBox1.Text != "OFF")
            {
                this.run1.PerformClick();
            }
            if (channelBox2.Text != "OFF")
            {
                this.run2.PerformClick();
            }
            if (channelBox3.Text != "OFF")
            {
                this.run3.PerformClick();
            }
            if (channelBox4.Text != "OFF")
            {
                this.run4.PerformClick();
            }
            if (channelBox5.Text != "OFF")
            {
                this.run5.PerformClick();
            }
            if (channelBox6.Text != "OFF")
            {
                this.run6.PerformClick();
            }
            if (channelBox7.Text != "OFF")
            {
                this.run7.PerformClick();
            }
            if (channelBox8.Text != "OFF")
            {
                this.run8.PerformClick();
            }

            serialnumberbox1.Focus();
        }







        public void RuntoolStripProgressBar(ToolStripProgressBar bar, int barStart, int barStop, int step)
        {
            bar.Value = barStart;

            for (int i=0; bar.Value < barStop; i++)
            {
                bar.Value = bar.Value + 1 * step;
            }

        }

        public void RunProgressBar(object number)
        {
            this.Invoke(new MethodInvoker(() => toolStripProgressBar1.Maximum = (int)number));
            

            for (int i=0; i<=(int)number;i++)
            {
                this.Invoke(new MethodInvoker(() => toolStripProgressBar1.Value = i));

                
                Application.DoEvents();
            }

        }

        public MultiMeter Initialize_U3606A(string connectstring)
        {
            
            if (connectstring != "")
            {
                MultiMeter dmm = new MultiMeter(connectstring);
                return dmm;
            }

            return null;
        }

        public Agilent Initialize_U2751A(string connectstring)
        {
            if (connectstring != "")
            {
                Agilent sw_A = new Agilent();
                sw_A.InitializeU2751A_WELLA(connectstring);
                
                return sw_A;
            }

            return null;
        }

        private bool SelectCH_SWA(string CH)
        {
            if (CH == "OFF")
            {
                channelClosed = true;
            }
            else if (CH == "CH1")
            {
                SW_A.SetRelayWellA(1, 0);
                channelClosed = false;
            }
            else if (CH == "CH2")
            {
                SW_A.SetRelayWellA(2, 0);
                channelClosed = false;
            }
            else if (CH == "CH3")
            {
                SW_A.SetRelayWellA(4, 0);
                channelClosed = false;
            }
            else if (CH == "CH4")
            {
                SW_A.SetRelayWellA(8, 0);
                channelClosed = false;
            }
            else if (CH == "CH5")
            {
                SW_A.SetRelayWellA(0, 1);
                channelClosed = false;
            }
            else if (CH == "CH6")
            {
                SW_A.SetRelayWellA(0, 2);
                channelClosed = false;
            }
            else if (CH == "CH7")
            {
                SW_A.SetRelayWellA(0, 4);
                channelClosed = false;
            }
            else if (CH == "CH8")
            {
                SW_A.SetRelayWellA(0, 8);
                channelClosed = false;
            }
            else
            {
                channelClosed = true;
            }



            return channelClosed;
        }

        public double RunTest(string CH, ToolStripProgressBar bar, Label testvallabel, Label testresultlabel, ToolStripStatusLabel progresslabel)
        {

            Thread.Sleep(DmmX8_Plus.Properties.Settings.Default.WaitInMS);


            double measuredvalue = 0.00;
            
            if ((devReady == false)||(CH.Equals("OFF")))
            {
                progresslabel.Text = "Closed";
                //MessageBox.Show("DMM or Switch is not ready!!! Please check the device in the Keysight IO Library.", "Instrument Not Available");
                return 0.00;
            }



            SelectCH_SWA(CH);

            int loops = DmmX8_Plus.Properties.Settings.Default.NumOfSamples;

            int progressbar_start = 40;

            //Thread myThread = new Thread(RunProgressBar);
            //myThread.IsBackground = true;
            
            //myThread.Start(loops);

            for (int i=0; i < loops; i++)
            {
                measuredvalue = DMM.MeasureChannelCurrent().average;

                measuredvalue = Math.Abs(measuredvalue);
                measuredvalue *= 1000;
                measuredvalue = Math.Round(measuredvalue, 5);

                progressbar_start += DmmX8_Plus.Properties.Settings.Default.progress_step;

                RuntoolStripProgressBar(bar, progressbar_start, 100, 1);

                testvallabel.Text = measuredvalue.ToString();

                if ((measuredvalue < DmmX8_Plus.Properties.Settings.Default.UpperLimitInMA) && (measuredvalue > DmmX8_Plus.Properties.Settings.Default.LowerLimitInMA))
                {
                    testresultlabel.Text = "Pass";
                    testresultlabel.ForeColor = Color.LawnGreen;
                    break;
                }
                else
                {
                    testresultlabel.Text = "Fail";
                    testresultlabel.ForeColor = Color.Red;
                    
                }

                Thread.Sleep(DmmX8_Plus.Properties.Settings.Default.IntervalOfSamplingInMS);
            }




            

            return measuredvalue;
        }

        private void serialnumberbox1_TextChanged(object sender, EventArgs e)
        {
            if (serialnumberbox1.TextLength >= SNLength)
            {
                if (channelBox2.Text != "OFF")
                {
                    serialnumberbox2.Focus();
                }
                else if (channelBox3.Text != "OFF")
                {
                    serialnumberbox3.Focus();
                }
                else if (channelBox4.Text != "OFF")
                {
                    serialnumberbox4.Focus();
                }
                else if (channelBox5.Text != "OFF")
                {
                    serialnumberbox5.Focus();
                }
                else if (channelBox6.Text != "OFF")
                {
                    serialnumberbox6.Focus();
                }
                else if (channelBox7.Text != "OFF")
                {
                    serialnumberbox7.Focus();
                }
                else if (channelBox8.Text != "OFF")
                {
                    serialnumberbox8.Focus();
                }

            }
        }

        private void serialnumberbox2_TextChanged(object sender, EventArgs e)
        {
            if (serialnumberbox2.TextLength >= SNLength)
            {
                if (channelBox3.Text != "OFF")
                {
                    serialnumberbox3.Focus();
                }
                else if (channelBox4.Text != "OFF")
                {
                    serialnumberbox4.Focus();
                }
                else if (channelBox5.Text != "OFF")
                {
                    serialnumberbox5.Focus();
                }
                else if (channelBox6.Text != "OFF")
                {
                    serialnumberbox6.Focus();
                }
                else if (channelBox7.Text != "OFF")
                {
                    serialnumberbox7.Focus();
                }
                else if (channelBox8.Text != "OFF")
                {
                    serialnumberbox8.Focus();
                }

            }
        }

        private void serialnumberbox3_TextChanged(object sender, EventArgs e)
        {
            if (serialnumberbox3.TextLength >= SNLength)
            {
                if (channelBox4.Text != "OFF")
                {
                    serialnumberbox4.Focus();
                }
                else if (channelBox5.Text != "OFF")
                {
                    serialnumberbox5.Focus();
                }
                else if (channelBox6.Text != "OFF")
                {
                    serialnumberbox6.Focus();
                }
                else if (channelBox7.Text != "OFF")
                {
                    serialnumberbox7.Focus();
                }
                else if (channelBox8.Text != "OFF")
                {
                    serialnumberbox8.Focus();
                }

            }
        }

        private void serialnumberbox4_TextChanged(object sender, EventArgs e)
        {
            if (serialnumberbox4.TextLength >= SNLength)
            {
                if (channelBox5.Text != "OFF")
                {
                    serialnumberbox5.Focus();
                }
                else if (channelBox6.Text != "OFF")
                {
                    serialnumberbox6.Focus();
                }
                else if (channelBox7.Text != "OFF")
                {
                    serialnumberbox7.Focus();
                }
                else if (channelBox8.Text != "OFF")
                {
                    serialnumberbox8.Focus();
                }

            }
        }

        private void serialnumberbox5_TextChanged(object sender, EventArgs e)
        {
            if (serialnumberbox5.TextLength >= SNLength)
            {
                if (channelBox6.Text != "OFF")
                {
                    serialnumberbox6.Focus();
                }
                else if (channelBox7.Text != "OFF")
                {
                    serialnumberbox7.Focus();
                }
                else if (channelBox8.Text != "OFF")
                {
                    serialnumberbox8.Focus();
                }

            }
        }

        private void serialnumberbox6_TextChanged(object sender, EventArgs e)
        {
            if (serialnumberbox6.TextLength >= SNLength)
            {
                if (channelBox7.Text != "OFF")
                {
                    serialnumberbox7.Focus();
                }
                else if (channelBox8.Text != "OFF")
                {
                    serialnumberbox8.Focus();
                }

            }
        }

        private void serialnumberbox7_TextChanged(object sender, EventArgs e)
        {
            if (serialnumberbox7.TextLength >= SNLength)
            {
                if (channelBox8.Text != "OFF")
                {
                    serialnumberbox8.Focus();
                }

            }
        }

        private void serialnumberbox8_TextChanged(object sender, EventArgs e)
        {
            if (serialnumberbox8.TextLength >= SNLength)
            {
                if (channelBox1.Text != "OFF")
                {
                    serialnumberbox1.Focus();
                }
                else if (channelBox2.Text != "OFF")
                {
                    serialnumberbox2.Focus();
                }
                else if (channelBox3.Text != "OFF")
                {
                    serialnumberbox3.Focus();
                }
                else if (channelBox4.Text != "OFF")
                {
                    serialnumberbox4.Focus();
                }
                else if (channelBox5.Text != "OFF")
                {
                    serialnumberbox5.Focus();
                }
                else if (channelBox6.Text != "OFF")
                {
                    serialnumberbox6.Focus();
                }
                else if (channelBox7.Text != "OFF")
                {
                    serialnumberbox7.Focus();
                }
                else
                {
                    serialnumberbox8.Focus();
                }

                allrun.Focus();
                allrun.PerformClick();

            }
        }
    }
}
