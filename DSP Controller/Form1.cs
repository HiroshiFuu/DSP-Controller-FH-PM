using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;
using System.Threading;
using System.Windows.Forms.DataVisualization.Charting;
using System.Collections;
using System.Text.RegularExpressions;

namespace DSP_Controller
{
    public partial class Form1 : Form
    {

        #region Members
        private SerialPort ComPort;
        private SerialMessage[] serialMsg;
        private SerialMessage serialMsgSend;
        private delegate void ReadDelegateHandler();
        private ReadDelegateHandler ReadDelegate;
        private delegate void ProcessDelegateHandler();
        private ProcessDelegateHandler ProcessDelegate;
        private delegate void SendDelegateHandler();
        private SendDelegateHandler SendDelegate;
        private delegate void SetTextDelegateHandler(string text);
        private SetTextDelegateHandler SetTextDelegate;
        private bool connected = false;
        private bool completed = false;
        private System.Windows.Forms.Timer timer;
        private System.Windows.Forms.Timer[] timerParas;
        private System.Timers.Timer timeout;
        private int[] plot_index;
        private uint command = 0x10000000;  //Init Speed Mode
        private TextBox[] tbParas;
        private UInt16 ParaChangedFlag = 0x0000;
        private byte[] ParaAddress = { 0x11, 0x12, 0x13, 0x31, 0x32, 0x33, 0x41, 0x42, 0x43 };
        private UInt16 ParaReceivedFlag = 0x01FF;
        private byte[] ReqAddress = { 0x10, 0x20, 0x30, 0x40, 0x50, 0x01, 0x02, 0x03, 0x04, 0x90, 0xA0 }; //Rv1:9, Rv2:10
        private string[] Paras;
        private int ReqIndex = 0;
        private ArrayList curveComm;
        private int stepCnt = 0;
        private int timerCnt = 1;
        private byte[] RevMessages = new byte[1024];
        private int RevMessagesOffset = 0;
        private Chart[] charts;
        private string[] seriesNames;
        private int ReadMsgIndex = 0;
        private int LoadMsgIndex = 0;
        private double[] initParas;
        #endregion

        #region From Init Function
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ComPort = new SerialPort();
            serialMsg = new SerialMessage[8];
            for (int i = 0; i < 8; i++)
                serialMsg[i] = new SerialMessage();
            serialMsgSend = new SerialMessage();

            //initial serial communication
            ComPort = new SerialPort();
            ComPort.BaudRate = 9600;
            ComPort.DataBits = 8;
            ComPort.Parity = Parity.None;
            ComPort.StopBits = StopBits.One;
            ComPort.DataReceived += new SerialDataReceivedEventHandler(serialPort_DataReceived);

            //initial timer for erroneous message
            timer = new System.Windows.Forms.Timer();
            timer.Interval = 300;
            timer.Tick += new EventHandler(timer_Tick);

            //initial timer for setting parameter
            timerParas = new System.Windows.Forms.Timer[11];
            for (int i = 0; i < 11; i++)
            {
                timerParas[i] = new System.Windows.Forms.Timer();
                timerParas[i].Interval = 100;
                timerParas[i].Tick += new EventHandler(timerParas_Tick);
            }

            //initial timer for timed-out display
            timeout = new System.Timers.Timer();
            timeout.Interval = 1200;
            timeout.AutoReset = false;
            timeout.Elapsed += new System.Timers.ElapsedEventHandler(OnTimedEvent);

            ReadDelegate = new ReadDelegateHandler(ReadMessages);
            ProcessDelegate = new ProcessDelegateHandler(ProcessMessages);
            SendDelegate = new SendDelegateHandler(SendMessages);
            SetTextDelegate = new SetTextDelegateHandler(SetText);

            lbTimeout.Text = "";
            plot_index = new int[15];
            Paras = new string[15];
            tbParas = new TextBox[9];
            tbParas[0] = tb_SKpSpeed;
            tbParas[1] = tb_SKiSpeed;
            tbParas[2] = tb_SSatSpeed;
            tbParas[3] = tb_idKpCurrent;
            tbParas[4] = tb_idKiCurrent;
            tbParas[5] = tb_idSatCurrent;
            tbParas[6] = tb_iqKpCurrent;
            tbParas[7] = tb_iqKiCurrent;
            tbParas[8] = tb_iqSatCurrent;
            initParas = new double[9];

            curveComm = new ArrayList();

            for (int i = 0; i < 200; i++)
            {
                chartSpeed.Series["speed"].Points.AddXY(i, 0);
                chartTorque.Series["torque"].Points.AddXY(i, 0);
                chart_id.Series["id"].Points.AddXY(i, 0);
                chart_iq.Series["iq"].Points.AddXY(i, 0);
                chartLd.Series["ld"].Points.AddXY(i, 0);
                chartLq.Series["lq"].Points.AddXY(i, 0);
                chartRs.Series["rs"].Points.AddXY(i, 0);
                chartPsi.Series["psi"].Points.AddXY(i, 0);
                chartRv1.Series["Rv1"].Points.AddXY(i, 0);
                chartRv2.Series["Rv2"].Points.AddXY(i, 0);
            }
            setChartStyle(chartSpeed, "speed");
            setChartStyle(chartTorque, "torque");
            setChartStyle(chart_id, "id");
            setChartStyle(chart_iq, "iq");
            setChartStyle(chartLd, "ld");
            setChartStyle(chartLq, "lq");
            setChartStyle(chartRs, "rs");
            setChartStyle(chartPsi, "psi");
            setChartStyle(chartRv1, "Rv1");
            setChartStyle(chartRv2, "Rv2");

            charts = new Chart[10];
            seriesNames = new string[10];
            charts[0] = chartSpeed;
            seriesNames[0] = "speed";
            charts[1] = chartTorque;
            seriesNames[1] = "torque";
            charts[2] = chart_id;
            seriesNames[2] = "id";
            charts[3] = chart_iq;
            seriesNames[3] = "iq";
            charts[4] = chartLd;
            seriesNames[4] = "ld";
            charts[6] = chartLq;
            seriesNames[6] = "lq";
            charts[5] = chartRs;
            seriesNames[5] = "rs";
            charts[7] = chartPsi;
            seriesNames[7] = "psi";
            charts[8] = chartRv1;
            seriesNames[8] = "Rv1";
            charts[9] = chartRv2;
            seriesNames[9] = "Rv2";

            refreshPorts();
        }

        private void setChartStyle(Chart chart, string series)
        {
            chart.ChartAreas[0].AxisX.Maximum = 200;
            chart.ChartAreas[0].AxisX.MajorGrid.LineDashStyle = ChartDashStyle.Dash;
            chart.ChartAreas[0].AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dash;
            chart.Series[series].BorderColor = Color.Blue;
            chart.Series[series].BorderWidth = 2;
            Series series2 = new Series();
            series2.Name = "Series2";
            series2.ChartType = SeriesChartType.Point;
            series2.Color = Color.DodgerBlue;
            series2.BorderColor = Color.Blue;
            series2.IsVisibleInLegend = false;
            series2.Points.AddXY(0, 0);
            chart.Series.Add(series2);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ComPort.IsOpen)
                ComPort.Close();
        }
        #endregion

        #region Step Point class
        class StepPoint
        {
            public int time;
            public float value;

            public StepPoint(int t, float v)
            {
                time = t;
                value = v;
            }
        }
        #endregion

        #region Methods
        private void refreshPorts()
        {
            if (cbCom.Enabled)
            {
                cbCom.SelectedIndex = -1;
                cbCom.Items.Clear();
            }
            foreach (string port in SerialPort.GetPortNames())
            {
                if (cbCom.Enabled)
                    cbCom.Items.Add(port);
            }
        }

        private void setControls(bool state)
        {
            btnApplySpeed.Enabled = state;
            btnApplySpeedCurve.Enabled = state;
            btnApply_iq.Enabled = state;
            btnApply_iqCurve.Enabled = state;
            btnSet.Enabled = state;
            btnResetPara.Enabled = state;
            cbCom.Enabled = !state;
            trackBarRun.Enabled = state;
        }

        private void WriteToFile(Chart chart, string seriesName)
        {
            if(!File.Exists("./Data"))
                Directory.CreateDirectory("./Data");
            if (!File.Exists("./Data/" + chart.Name.Replace("chart", "").Replace("_", "")))
                Directory.CreateDirectory("./Data/" + chart.Name.Replace("chart", "").Replace("_", ""));
            StreamWriter file = new StreamWriter("./Data/" + chart.Name.Replace("chart", "").Replace("_", "") + "/" + seriesName + "_" + string.Format("{0:yyyy-MM-dd_HH-mm-ss}", DateTime.Now) + ".csv", false);
            for (int i = 0; i < 200; i++)
                file.Write(chart.Series[seriesName].Points[i].YValues[0].ToString() + ",");
            file.Flush();
            file.Close();
        }

        private void checkParaFilled()
        {
            bool flag = false;
            for (byte i = 0; i < 9; i++)
                if (tbParas[i].Text.Trim() != "")
                    flag = true;
            for (byte i = 0; i < 9; i++)
                if (tbParas[i].Text.Trim() == "-")
                    flag = false;
            if (connected)
                btnSet.Enabled = flag;
        }

        private void validatePara()
        {
            bool flag = true;
            for (byte i = 0; i < 9; i++)
                if (tbParas[i].Text.Trim() != "")
                    try
                    {
                        float.Parse(tbParas[i].Text.Trim());
                    }
                    catch
                    {
                        flag = false;
                    }
            if (connected)
                btnSet.Enabled = flag;
        }

        private bool validateCurveEntry(RichTextBox rtb)
        {
            curveComm.Clear();
            stepCnt = 0;
            timerCnt = 1;
            if (rtb.Lines.Length == 0)
                return false;
            foreach (string line in rtb.Lines)
            {
                if (line.Trim() == "")
                    return false;
                string[] xy = line.Split(',');
                if (!Regex.Match(line.Trim(), @"\d+,-?\d+.?\d*").Success)
                    return false;
                curveComm.Add(new StepPoint(Int32.Parse(xy[0]), float.Parse(xy[1])));
            }
            return true;
        }

        private void SendConnectionComm(Byte StartId, Byte MsgId)
        {
            serialMsgSend.MsgId = MsgId;
            serialMsgSend.MsgData = BitConverter.GetBytes(0);
            serialMsgSend.setMsg(StartId);
            SendMessages();
            if (ComPort.IsOpen)
                ComPort.Write(serialMsgSend.Message, 0, 8);
            else
            {
                try
                {
                    ComPort.Close();
                }
                catch
                {
                    MessageBox.Show("Connection Lost!");
                }
            }
        }

        private void SendCommand()
        {
            serialMsgSend.MsgId = 0x01;
            serialMsgSend.MsgData = BitConverter.GetBytes(command);
            serialMsgSend.setMsg();
            SendMessages();
            if (ComPort.IsOpen)
                ComPort.Write(serialMsgSend.Message, 0, 8);
            else
            {
                try
                {
                    ComPort.Close();
                }
                catch
                {
                    MessageBox.Show("Connection Lost!");
                }
            }
        }

        private void SendPara(byte Addr)
        {
            serialMsgSend.MsgId = Addr;
            serialMsgSend.setMsg();
            SendMessages();
            if (ComPort.IsOpen)
                ComPort.Write(serialMsgSend.Message, 0, 8);
            else
            {
                try
                {
                    ComPort.Close();
                }
                catch
                {
                    MessageBox.Show("Connection Lost!");
                }
            }
        }

        private void SendMessage(byte Addr, float value)
        {
            SendMessage(Addr, value.ToString());
        }

        private void SendMessage(byte Addr, string value)
        {
            serialMsgSend.MsgId = Addr;
            serialMsgSend.MsgData = BitConverter.GetBytes(float.Parse(value));
            serialMsgSend.setMsg();
            SendMessages();
            if (ComPort.IsOpen)
                ComPort.Write(serialMsgSend.Message, 0, 8);
            else
            {
                try
                {
                    ComPort.Close();
                }
                catch
                {
                    MessageBox.Show("Connection Lost!");
                }
            }
        }

        private int IncReq()
        {
            int number = ReqAddress.Length - 2;
            if (cbRv1Update.Checked)
                number = 10;
            if (cbRv2Update.Checked)
                number = 11;
            if (ReqIndex == number)
                ReqIndex = 0;
            if (!cbRv1Update.Checked && cbRv2Update.Checked && ReqIndex == 8)
                ReqIndex = 10;
            return ReqIndex++;
        }

        private void SetText(string text)
        {
            lbTimeout.Text = "";
            lbTimeout.Text += text;
            lbTimeout.Visible = true;
        }
		
        private void ResetChart(int index)
        {
            Chart chart = charts[index];
            for (int i = 0; i < 200; i++)
            {
                chart.Series[seriesNames[index]].Points[i].SetValueY(0);
            }
            plot_index[index] = 0;
        }
        #endregion

        #region Timer Event
        private void timer_Tick(object sender, EventArgs e)
        {
            try
            {
                ComPort.DiscardInBuffer();
                ComPort.ReadExisting();
                RevMessages = new byte[1024];
                RevMessagesOffset = 0;
            }
            catch
            { }
        }

        private void OnTimedEvent(object source, System.Timers.ElapsedEventArgs e)
        {
            this.Invoke(new EventHandler(hideTimeroutVisible));
        }

        private void hideTimeroutVisible(object sender, EventArgs e)
        {
            lbTimeout.Text = "";
            lbTimeout.Visible = false;
        }

        private void timerCommand_Tick(object sender, EventArgs e)
        {
            if (ComPort.IsOpen)
            {
                serialMsgSend.MsgId = ReqAddress[IncReq()];
                serialMsgSend.MsgData = BitConverter.GetBytes(0);
                serialMsgSend.MsgData[3] = 0x11;
                serialMsgSend.setMsg(0x1A);
                //SendMessages();
                try
                {
                    ComPort.Write(serialMsgSend.Message, 0, 8);
                }
                catch
                {
                    btnConnect_Click(null, null);
                    Invoke(SetTextDelegate, "DSP Timeout\n");
                    timer.Enabled = false;
                    if (timeout.Enabled)
                        timeout.Stop();
                    timeout.Start();
                    return;
                }
            }
        }

        private void timerCurve_Tick(object sender, EventArgs e)
        {
            StepPoint step = (StepPoint)curveComm[stepCnt];
            if (step.time == timerCnt++)
            {
                if (!rtbSpeed.Enabled)
                    SendMessage(0x10, step.value);
                if (!rtb_iq.Enabled)
                    SendMessage(0x20, step.value);
                stepCnt++;
                //MessageBox.Show(step.value.ToString());
                if (stepCnt == curveComm.Count)
                {
                    timerCurve.Stop();
                    btnApplySpeedCurve.Enabled = btnApply_iqCurve.Enabled = true;
                    rtbSpeed.Enabled = rtb_iq.Enabled = true;
                    timerCnt = 1;
                    curveComm.Clear();
                }
            }
        }

        private void timerCheckCom_Tick(object sender, EventArgs e)
        {
            refreshPorts();
        }

        private void timerConnect_Tick(object sender, EventArgs e)
        {
            SendConnectionComm(0x1C, 0x00);
        }

        private void timerParas_Tick(object sender, EventArgs e)
        {
            System.Windows.Forms.Timer timer = (System.Windows.Forms.Timer)sender;
            int index = Array.IndexOf(timerParas, timer);
            serialMsgSend.MsgData = BitConverter.GetBytes(float.Parse(tbParas[index].Text));
            SendPara(ParaAddress[index]);
            Thread.Sleep(20);
        }

        private void timerConnected_Tick(object sender, EventArgs e)
        {
            SendConnectionComm(0x1C, 0x00);
        }
        #endregion

        #region Control EventHandler
        private void cbCom_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbCom.SelectedIndex != -1)
            {
                ComPort.PortName = cbCom.SelectedItem.ToString();
                btnConnect.Enabled = true;
            }
            else
                btnConnect.Enabled = false;
        }

        private void cbCom_KeyDown(object sender, KeyEventArgs e)
        {
            if (cbCom.SelectedIndex != -1)
                if (e.KeyValue == 13)
                    btnConnect_Click(btnConnect, null);
        }

        private void tb_TextChanged(object sender, EventArgs e)
        {
            TextBox tb = (TextBox)sender;
            if (tb.Text.Trim() == "-" || tb.Text.Trim() == "")
            {
                checkParaFilled();
                return;
            }
            try
            {
                errorProvider1.Clear();
                tb.Text = tb.Text.Trim();
                float.Parse(tb.Text);
                validatePara();
                UInt16 mask = 0x0001;
                for (byte i = 0; i < 9; i++)
                {
                    if (sender == tbParas[i])
                        ParaChangedFlag |= mask;
                    mask <<= 0x1;
                    Paras[i] = tbParas[i].Text;
                }
            }
            catch
            {
                errorProvider1.SetError(tb, "Invalid Input");
                btnSet.Enabled = false;
                tb.Focus();
                tb.SelectAll();
            }
        }

        private void rb_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton rb = (RadioButton)sender;
            command &= 0x00FFFFFF;
            if (rb.Tag.ToString() == "1")
            {
                command |= 0x10000000;
                lb_iq.Visible = tb_iq.Visible = btnApply_iq.Visible = lb_iqCurve.Enabled = rtb_iq.Enabled = btnApply_iqCurve.Enabled = false;
                return;
            }
            if (rb.Tag.ToString() == "2")
            {
                command |= 0x01000000;
                lb_iq.Visible = tb_iq.Visible = btnApply_iq.Visible = lb_iqCurve.Enabled = rtb_iq.Enabled = btnApply_iqCurve.Enabled = true;
                return;
            }
        }

        private void rtbLog_TextChanged(object sender, EventArgs e)
        {
            RichTextBox rtb = (RichTextBox)sender;
            if (rtb.Lines.Count() > 400)
                rtb.Text = "";
        }

        private void trackBarRun_Scroll(object sender, EventArgs e)
        {
            if (trackBarRun.Value == 0)
            {
                picStatus.Image = Properties.Resources.red;
                picStatus.Refresh();
                command &= 0xFFFFFF00;

                SendCommand();
            }
            if (trackBarRun.Value == 1)
            {
                picStatus.Image = Properties.Resources.green;
                picStatus.Refresh();
                command &= 0xFFFFFF00;
                command |= 0x00000001;
                SendCommand();
            }
            btnConnect.Focus();
        }

        private void cbAuto_CheckedChanged(object sender, EventArgs e)
        {
            cbSpeed.Checked = cbTorque.Checked = cb_id.Checked = cb_iq.Checked = cbLm.Checked = cbRs.Checked = cbLls.Checked = cbRr.Checked = cbAuto.Checked;
            cbSpeed.Enabled = cbTorque.Enabled = cb_id.Enabled = cb_iq.Enabled = cbLm.Enabled = cbRs.Enabled = cbLls.Enabled = cbRr.Enabled = !cbAuto.Checked;
            if (cbRv1Update.Checked)
                cbRv1Save.Checked = cbAuto.Checked;
            if (cbRv2Update.Checked)
                cbRv2Save.Checked = cbAuto.Checked;
        }

        private void cbRv1Update_CheckedChanged(object sender, EventArgs e)
        {
            cbRv1Save.Enabled = cbRv1Update.Checked;
            cbRv1Save.Checked = cbAuto.Checked;
        }

        private void cbRv2Update_CheckedChanged(object sender, EventArgs e)
        {
            cbRv2Save.Enabled = cbRv2Update.Checked;
            cbRv2Save.Checked = cbAuto.Checked;
        }

        private void tbConstentSpeed_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (btnApplySpeed.Enabled && e.KeyChar == 13)
                btnApplySpeed_Click(null, null);
        }

        private void tbConstentTorque_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (btnApply_iq.Enabled && e.KeyChar == 13)
                btnApply_iq_Click(null, null);
        }

        private void tb_Paras_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (btnSet.Enabled && e.KeyChar == 13)
                btnSet_Click(null, null);
        }

        private void rtbLog_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            RichTextBox rtb = (RichTextBox)sender;
            rtb.Text = "";
        }
        #endregion

        #region Message Event
        private void serialPort_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            if (ComPort.BytesToRead < 1)
                return;
            try
            {
                int toRead = ComPort.BytesToRead;
                ComPort.Read(RevMessages, RevMessagesOffset, toRead);
                RevMessagesOffset += toRead;
            }
            catch { }
            BeginInvoke(ReadDelegate);
        }

        private void ReadMessages()
        {
            if (RevMessagesOffset >= 8)
            {
                int i = 0;
                for (; i <= RevMessagesOffset - 8; i++)
                {
                    if (RevMessages[i] == 0x1A || RevMessages[i] == 0x17 || RevMessages[i] == 0x1B)
                    {
                        if (RevMessages[i + 1] == 0x1A)
                            continue;
                        //timer.Stop();
                        //timer.Start();
                        for (int j = 0; j < 8; j++)
                            serialMsg[ReadMsgIndex].Message[j] = RevMessages[i + j];
                        serialMsg[ReadMsgIndex].getMsg();
                        if (serialMsg[ReadMsgIndex].chkSumErr)
                        {
                            try
                            {
                                BeginInvoke(SetTextDelegate, "0x" + serialMsg[ReadMsgIndex].StartId.ToString("X").PadLeft(2, '0') + serialMsg[ReadMsgIndex].MsgId.ToString("X").PadLeft(2, '0') + " CRC Error\n");
                            }
                            catch { }
                            if (timeout.Enabled)
                                timeout.Stop();
                            timeout.Start();
                        }
                        if (RevMessages[i] != 0x1A)
                        {
                            string DisplayRx = "";
                            DisplayRx = BitConverter.ToString(serialMsg[ReadMsgIndex].Message);
                            rtbRx.SelectedText = string.Empty;
                            rtbRx.AppendText("Rev " + DisplayRx + "\n");
                            rtbRx.ScrollToCaret();
                        }
                        BeginInvoke(ProcessDelegate);
                        ReadMsgIndex++;
                        if (ReadMsgIndex == 8)
                            ReadMsgIndex = 0;
                        i += 7;
                    }
                }
                i--;
                for (int j = 0; j < i; j++)
                    RevMessages[j] = RevMessages[j + i];
                RevMessagesOffset -= i;
            }
        }

        private void ProcessMessages()
        {
            if (!serialMsg[LoadMsgIndex].chkSumErr)
                if (serialMsg[LoadMsgIndex].StartId == 0x1A)
                {
                    switch (serialMsg[LoadMsgIndex].MsgId)
                    {
                        case 0x10:
                            tbR_Speed.Text = BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0).ToString();
                            chartSpeed.Series["speed"].Points[plot_index[0]].SetValueY(BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chartSpeed.Series["Series2"].Points[0].SetValueXY(plot_index[0]++, BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chartSpeed.ChartAreas[0].RecalculateAxesScale();
                            if (plot_index[0] == 200)
                            {
                                plot_index[0] = 0;
                                if (cbSpeed.Checked)
                                    WriteToFile(chartSpeed, "speed");
                            }
                            break;
                        case 0x20:
                            tbR_Torque.Text = BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0).ToString();
                            chartTorque.Series["torque"].Points[plot_index[1]].SetValueY(BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chartTorque.Series["Series2"].Points[0].SetValueXY(plot_index[1]++, BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chartTorque.ChartAreas[0].RecalculateAxesScale();
                            if (plot_index[1] == 200)
                            {
                                plot_index[1] = 0;
                                if (cbTorque.Checked)
                                    WriteToFile(chartTorque, "torque");
                            }
                            break;
                        case 0x30:
                            tbR_id.Text = BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0).ToString();
                            chart_id.Series["id"].Points[plot_index[2]].SetValueY(BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chart_id.Series["Series2"].Points[0].SetValueXY(plot_index[2]++, BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chart_id.ChartAreas[0].RecalculateAxesScale();
                            if (plot_index[2] == 200)
                            {
                                plot_index[2] = 0;
                                if (cb_id.Checked)
                                    WriteToFile(chart_id, "id");
                            }
                            break;
                        case 0x40:
                            tbR_iq.Text = BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0).ToString();
                            chart_iq.Series["iq"].Points[plot_index[3]].SetValueY(BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chart_iq.Series["Series2"].Points[0].SetValueXY(plot_index[3]++, BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chart_iq.ChartAreas[0].RecalculateAxesScale();
                            if (plot_index[3] == 200)
                            {
                                plot_index[3] = 0;
                                if (cb_iq.Checked)
                                    WriteToFile(chart_iq, "iq");
                            }
                            break;
                        case 0x50:
                            tbR_DC.Text = BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0).ToString();
                            break;
                        case 0x01:
                            tbLd.Text = BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0).ToString();
                            chartLd.Series["ld"].Points[plot_index[4]].SetValueY(BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chartLd.Series["Series2"].Points[0].SetValueXY(plot_index[4]++, BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chartLd.ChartAreas[0].RecalculateAxesScale();
                            if (plot_index[4] == 200)
                            {
                                plot_index[4] = 0;
                                if (cbLm.Checked)
                                    WriteToFile(chartLd, "ld");
                            }
                            break;
                        case 0x02:
                            tbLq.Text = BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0).ToString();
                            chartLq.Series["lq"].Points[plot_index[5]].SetValueY(BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chartLq.Series["Series2"].Points[0].SetValueXY(plot_index[5]++, BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chartLq.ChartAreas[0].RecalculateAxesScale();
                            if (plot_index[5] == 200)
                            {
                                plot_index[5] = 0;
                                if (cbLls.Checked)
                                    WriteToFile(chartLq, "lq");
                            }
                            break;
                        case 0x03:
                            tbRs.Text = BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0).ToString();
                            chartRs.Series["rs"].Points[plot_index[6]].SetValueY(BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chartRs.Series["Series2"].Points[0].SetValueXY(plot_index[6]++, BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chartRs.ChartAreas[0].RecalculateAxesScale();
                            if (plot_index[6] == 200)
                            {
                                plot_index[6] = 0;
                                if (cbRs.Checked)
                                    WriteToFile(chartRs, "rs");
                            }
                            break;
                        case 0x04:
                            tbPsi.Text = BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0).ToString();
                            chartPsi.Series["psi"].Points[plot_index[7]].SetValueY(BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chartPsi.Series["Series2"].Points[0].SetValueXY(plot_index[7]++, BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chartPsi.ChartAreas[0].RecalculateAxesScale();
                            if (plot_index[7] == 200)
                            {
                                plot_index[7] = 0;
                                if (cbRr.Checked)
                                    WriteToFile(chartPsi, "psi");
                            }
                            break;
                        case 0x90:
                            chartRv1.Series["Rv1"].Points[plot_index[8]].SetValueY(BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chartRv1.Series["Series2"].Points[0].SetValueXY(plot_index[8]++, BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chartRv1.ChartAreas[0].RecalculateAxesScale();
                            if (plot_index[8] == 200)
                            {
                                plot_index[8] = 0;
                                if (cbRv1Save.Checked)
                                    WriteToFile(chartRv1, "Rv1");
                            }
                            break;
                        case 0xA0:
                            chartRv2.Series["Rv2"].Points[plot_index[9]].SetValueY(BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chartRv2.Series["Series2"].Points[0].SetValueXY(plot_index[9]++, BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0));
                            chartRv2.ChartAreas[0].RecalculateAxesScale();
                            if (plot_index[9] == 200)
                            {
                                plot_index[9] = 0;
                                if (cbRv2Save.Checked)
                                    WriteToFile(chartRv2, "Rv2");
                            }
                            break;
                    }
                }
                else if (serialMsg[LoadMsgIndex].StartId == 0x1B)
                {
                    int target = Array.IndexOf(ParaAddress, serialMsg[LoadMsgIndex].MsgId);
                    if (target < 0)
                        return;
                    timerParas[target].Stop();
                    //switch (serialMsg[LoadMsgIndex].MsgId)
                    //{
                    //    case 0x01:
                    //        SendCommand();
                    //        break;
                    //    case 0x10:
                    //    case 0x20:
                    //        if ((command & 0x20) != 0)
                    //            SendMessage(0x10, tbSpeed.Text);
                    //        if ((command & 0x80) != 0)
                    //            SendMessage(0x20, tbTorque.Text);
                    //        if ((command & 0x10) != 0 || (command & 0x04) != 0)
                    //        {
                    //            if (timerCurve.Enabled)
                    //            {
                    //                timerCurve.Stop();
                    //                timerCurve.Start();
                    //            }
                    //            StepPoint step = (StepPoint)curveComm[stepCnt - 1];
                    //            if ((command & 0x10) != 0)
                    //                SendMessage(0x10, step.value);
                    //            if ((command & 0x04) != 0)
                    //                SendMessage(0x20, step.value);
                    //        }
                    //        break;
                    //    case 0x30:
                    //        if ((command & 0x20) != 0 || (command & 0x08) != 0)
                    //            SendMessage(0x30, tb_id.Text);
                    //        if ((command & 0x10) != 0 || (command & 0x04) != 0)
                    //            SendMessage(0x30, tb_id2.Text);
                    //        break;
                    //    default:
                    //        int index = Array.IndexOf(ParaAddress, serialMsg[LoadMsgIndex].MsgId);
                    //        SendMessage(ParaAddress[index], Paras[index]);
                    //        break;
                    //
                }
                else if (serialMsg[LoadMsgIndex].StartId == 0x17)
                {
                    timerConnect.Stop();
                    UInt16 mask = 0x1;
                    int target = Array.IndexOf(ParaAddress, serialMsg[LoadMsgIndex].MsgId);
                    if (target < 0)
                        return;
                    for (int i = 0; i < target; i++)
                        mask <<= 1;
                    SendConnectionComm(0x27, serialMsg[LoadMsgIndex].MsgId);
                    if (tbParas[target].Text != "")
                        return;
                    tbParas[target].Text = BitConverter.ToSingle(serialMsg[LoadMsgIndex].MsgData, 0).ToString();
                    initParas[target] = BitConverter.ToDouble(serialMsg[LoadMsgIndex].MsgData, 0);
                    ParaReceivedFlag ^= mask;
                    if (ParaReceivedFlag == 0x0)
                    {
                        timerConnected.Stop();
                        completed = true;
                        setControls(true);
                        timer.Enabled = true;
                        SendMessage(0x30, tb_id.Text);
                        Thread.Sleep(20);
                        SendMessage(0x10, tbSpeed.Text);
                        Thread.Sleep(20);
                        SendMessage(0x10, tb_iq.Text);
                        Thread.Sleep(20);
                        SendCommand();
                        timerCommand.Start();
                    }
                }
            LoadMsgIndex++;
            if (LoadMsgIndex == 8)
                LoadMsgIndex = 0;
        }

        private void SendMessages()
        {
            string DisplayTx = "Tx ";
            DisplayTx += BitConverter.ToString(serialMsgSend.Message);
            rtbTx.SelectedText = string.Empty;
            rtbTx.AppendText(DisplayTx + "\n");
            rtbTx.ScrollToCaret();
        }
        #endregion

        #region Button EventHandler
        private void btnConnect_Click(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            if (!ComPort.IsOpen)
            {
                try
                {
                    ComPort.Open();
                    connected = true;
                    timerCheckCom.Enabled = false;
                    btn.Text = "Disconnect";
                    timerConnect.Start();
                    timerConnected.Start();
                    //setControls(true);
                    //timer.Enabled = true;
                    //timerCommand.Start();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
            else
            {
                try
                {
                    timerCommand.Stop();
                    timerCurve.Stop();
                    timeout.Stop();
                    timerConnect.Stop();
                    timerConnected.Stop();
                    for (int i = 0; i < 9; i++)
                    {
                        tbParas[i].Text = "";
                        timerParas[i].Stop();
                    }
                    RevMessages = new byte[1024];
                    RevMessagesOffset = 0;
                    ReadMsgIndex = 0;
                    LoadMsgIndex = 0;
                    ParaReceivedFlag = 0x01FF;
                    lbTimeout.Visible = false;
                    ComPort.Close();
                    btn.Text = "Connect";
                    connected = completed = false;
                    trackBarRun.Value = 0;
                    picStatus.Image = Properties.Resources.red;
                    picStatus.Refresh();
                    rbSpeed.Checked = true;
                    tb_id.Text = "0";
                    tb_id2.Text = "0";
                    tb_iq.Text = "0";
                    tbSpeed.Text = "0";
                    command = 0x10000000;
                    setControls(false);
                    timerCheckCom.Start();
                    refreshPorts();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }
            }
        }

        private void btnSet_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 9; i++)
            {
                if ((ParaChangedFlag & 0x1) == 1)
                {
                    serialMsgSend.MsgData = BitConverter.GetBytes(float.Parse(tbParas[i].Text));
                    SendPara(ParaAddress[i]);
                    timerParas[i].Start();
                    Thread.Sleep(20);
                }
                ParaChangedFlag >>= 0x1;
            }
        }

        private void btnShow_Click(object sender, EventArgs e)
        {
            if (btnShow.Text == "Show Message")
            {
                btnShow.Text = "Hide Message";
                lbMessage.Visible = true;
                rtbRx.Visible = true;
            }
            else
            {
                btnShow.Text = "Show Message";
                lbMessage.Visible = false;
                rtbRx.Visible = false;
            }
        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            serialMsgSend.StartId = Convert.ToByte(tbCommand.Text.Substring(0, 2), 16);
            serialMsgSend.MsgId = Convert.ToByte(tbCommandId.Text.Substring(0, 2), 16);
            serialMsgSend.MsgData[0] = Convert.ToByte(tbCommandData.Text.Substring(0, 2), 16);
            serialMsgSend.MsgData[1] = Convert.ToByte(tbCommandData.Text.Substring(2, 2), 16);
            serialMsgSend.MsgData[2] = Convert.ToByte(tbCommandData.Text.Substring(4, 2), 16);
            serialMsgSend.MsgData[3] = Convert.ToByte(tbCommandData.Text.Substring(6, 2), 16);
            serialMsgSend.setMsg(serialMsgSend.StartId);
            //if (ComPort.IsOpen)
            //    ComPort.Write(serialMsg[LoadMsgIndex].Message, 0, 8);
            //this.Invoke(SendDelegate);
            //serialMsg[LoadMsgIndex].Message[0] = Convert.ToByte(tbCommand.Text.Substring(0, 2), 16);
            //serialMsg[LoadMsgIndex].Message[1] = Convert.ToByte(tbCommandId.Text.Substring(0, 2), 16);
            //serialMsg[LoadMsgIndex].Message[5] = Convert.ToByte(tbCommandData.Text.Substring(0, 2), 16);
            //serialMsg[LoadMsgIndex].Message[4] = Convert.ToByte(tbCommandData.Text.Substring(2, 2), 16);
            //serialMsg[LoadMsgIndex].Message[3] = Convert.ToByte(tbCommandData.Text.Substring(4, 2), 16);
            //serialMsg[LoadMsgIndex].Message[2] = Convert.ToByte(tbCommandData.Text.Substring(6, 2), 16);
            //serialMsg[LoadMsgIndex].Message[6] = serialMsgSend.MsgChksum[0];
            //serialMsg[LoadMsgIndex].Message[7] = serialMsgSend.MsgChksum[1];
            //serialMsg[LoadMsgIndex].getMsg();
            //this.Invoke(ReadDelegate);
            //MessageBox.Show(serialMsg[LoadMsgIndex].chkSumErr.ToString());
        }

        private void btnApplySpeed_Click(object sender, EventArgs e)
        {
            errorProvider2.Clear();
            try
            {
                float.Parse(tbSpeed.Text);
                float.Parse(tb_id.Text);
            }
            catch
            {
                try
                {
                    float.Parse(tbSpeed.Text);
                }
                catch
                {
                    string value = tbSpeed.Text.Trim();
                    if (value != "" && value != "-")
                        errorProvider2.SetError(tbSpeed, "Invalid Value");
                }
                try
                {
                    float.Parse(tb_id.Text);
                }
                catch
                {
                    string value = tb_id.Text.Trim();
                    if (value != "" && value != "-")
                        errorProvider2.SetError(tb_id, "Invalid Value");
                }
                return;
            }
            command &= 0xFF00FFFF;
            command |= 0x00100000;
            SendCommand();
            Thread.Sleep(20);
            SendMessage(0x30, tb_id.Text);
            Thread.Sleep(20);
            SendMessage(0x10, tbSpeed.Text);
        }

        private void btnApply_iq_Click(object sender, EventArgs e)
        {
            errorProvider2.Clear();
            try
            {
                float.Parse(tbSpeed.Text);
                float.Parse(tb_id.Text);
            }
            catch
            {
                try
                {
                    float.Parse(tbSpeed.Text);
                }
                catch
                {
                    string value = tbSpeed.Text.Trim();
                    if (value != "" && value != "-")
                        errorProvider2.SetError(tbSpeed, "Invalid Value");
                }
                try
                {
                    float.Parse(tb_id.Text);
                }
                catch
                {
                    string value = tb_id.Text.Trim();
                    if (value != "" && value != "-")
                        errorProvider2.SetError(tb_id, "Invalid Value");
                }
                return;
            }
            command &= 0xFF00FFFF;
            command |= 0x00100000;
            SendCommand();
            Thread.Sleep(20);
            SendMessage(0x30, tb_id.Text);
            Thread.Sleep(20);
            SendMessage(0x20, tb_iq.Text);
        }

        private void btnApplySpeedCurve_Click(object sender, EventArgs e)
        {
            if (!validateCurveEntry(rtbSpeed))
            {
                errorProvider2.Clear();
                errorProvider2.SetError(lbSpeedCurve, "Error Entry");
                curveComm.Clear();
                stepCnt = 0;
                timerCnt = 1;
                rtbSpeed.SelectAll();
                rtbSpeed.ScrollToCaret();
                rtbSpeed.Focus();
                return;
            }
            btnApplySpeedCurve.Enabled = btnApply_iqCurve.Enabled = false;
            rtbSpeed.Enabled = false;
            command &= 0xFF00FFFF;
            command |= 0x00010000;
            SendCommand();
            Thread.Sleep(20);
            SendMessage(0x30, tb_id2.Text);
            timerCurve.Start();
        }

        private void btnApply_iqCurve_Click(object sender, EventArgs e)
        {
            if (!validateCurveEntry(rtb_iq))
            {
                errorProvider2.Clear();
                errorProvider2.SetError(lb_iqCurve, "Error Entry");
                curveComm.Clear();
                stepCnt = 0;
                timerCnt = 1;
                rtb_iq.SelectAll();
                rtb_iq.ScrollToCaret();
                rtb_iq.Focus();
                return;
            }
            btnApplySpeedCurve.Enabled = btnApply_iqCurve.Enabled = false;
            rtb_iq.Enabled = false;
            command &= 0xFFFF00FF;
            command |= 0x00000100;
            SendCommand();
            Thread.Sleep(20);
            SendMessage(0x30, tb_id2.Text);
            timerCurve.Start();
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            if (btn.Name.Length == 10)
                ResetChart(Int32.Parse(btn.Name[8].ToString() + btn.Name[9].ToString()) - 1);
            else
                ResetChart(Int32.Parse(btn.Name[8].ToString()) - 1);
        }

        private void btnResetPara_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 9; i++)
                tbParas[i].Text = initParas[i].ToString();
        }
        #endregion

    }
}
