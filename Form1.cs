using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.IO;
using System.IO.Ports;
using System.Threading;
using System.Drawing.Printing;
using System.Diagnostics;
using System.Runtime.InteropServices; // win32api 선언
using System.Linq;
using System.Xml;


namespace Kiosk
{
    /// <summary>
    /// public partial class Kiosk : Form
    /// </summary>
    public partial class Kiosk : Form
    {
        #region 전역변수 및 전역함수 선언
        string sensor_type;
        double double_cel;
        //시리얼 수신 parsing 부분 변수 선언
        double cel = 0, cel1 = 0; //온도계산
        byte[] readMsg = new byte[50];
        double datavalue; //데이터값 
        double datavalue3; //데이터값 
        string datavalue1 = string.Empty;
        string datatime = string.Empty; //현재시간 체크

        string node_id = string.Empty;
        static string msg_Data = string.Empty;
        string comp_Data = string.Empty;
        string msg_bufferData1; //임시기억장소 데이터값 
        string msg_bufferData2; //임시기억장소 데이터값
        string msg_bufferData3; //임시기억장소 데이터값

        string[] arr_FileSave = new string[10]; // 파일 저장 값

        private const byte STX = 0x7E;
        private const byte ETX = 0x7E;
        private const byte unUsed = 0x42; //7E 42 으로 시작하는 전문
        int read_msg_count = 0;
        bool chkFlagData = false; //
        bool chkFlagStx = false; // 첫번째 7E인지 판단 true 첫번째 아님 false 첫번째
        bool chkFlagNeed = true; //필요전문인지 아닌지 판단 true 필요 false 불필요

        string Backimg = @"\"; // 폼 스킨 
        string MonitorMode;
        string portselect = "COM1";
        string BautRateselect = "57600";
        string screensaver; 
        int iscreensaver; // 절전 딜레이 숫자로 치환 변수
        bool bscreen = true;     // 데이터 확인 flag
        int screenright;
        int screentop;
        int data_count = 0;
       
        //private Point mouseOffset;
        //private bool isMouseDown = false;

        // 모니터 ON/OFF API
        const int WM_SYSCOMMAND = 0x0112;
        const int SC_MONITORPOWER = 0xF170;
        const int MONITOR_ON = -1;
        const int MONITOR_OFF = 2;
        const int MONITOR_STANBY = 1;      
        
        [DllImport("user32.dll")]
        private static extern int SendMessage(int hWnd, int hMsg, int wParam, int lParam);   

        //수신 관련 
        private delegate void SetTextCallback(string strMessage);
        private delegate void SetMonitorCallback(int node);
        private const int MSG_SIZE = 75;
        private string messageBuffer = string.Empty;
        #endregion

        #region 생성자 선언
        /// <summary>
        /// 초기화
        /// </summary>
        public Kiosk()
        {
            InitializeComponent();
        }
        #endregion

        #region RS232 OPEN 메서드 
        /// <summary>
        /// 포트 접속
        /// </summary>
        private void RS232_Open()
        {
            /* RS232 통신 */
            if (!RS232.IsOpen)
            {

                try
                {
                    RS232.Encoding = System.Text.Encoding.GetEncoding(1252);
                    RS232.Open();
                }
                catch (Exception e)
                {
                    MessageBox.Show("msg -> " + e.StackTrace);
                }
            }
        }
        #endregion

        #region RS232 CLOSE 메서드
        /// <summary>
        /// 포트 접속 해제
        /// </summary>
        private void RS232_Close()
        {
            if (RS232.IsOpen)
            {
                RS232.Close();
            }
        }
        #endregion

        #region RS232 DataRecevie 메서드
        /// <summary>
        /// 데이터가 있을때
        /// </summary>
        private void RS232_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            string strMessage = RS232.ReadExisting();
            if (strMessage != string.Empty)
                messageparsing(strMessage);
        }
        #endregion

        #region 인코딩 메서드
        /// <summary>
        /// 바이트 배열로 인코딩
        /// </summary>
        /// <param name="strMessage">수신 전문</param>
        private byte[] DataEncoding(string strMessage)
        {
            return System.Text.Encoding.GetEncoding(1252).GetBytes(strMessage);
        }
        #endregion

        #region 메시지파싱 메서드
        /// <summary>
        /// 메세지 파싱
        /// </summary>
        /// <param name="strMessage"></param>
        private void messageparsing(string strMessage)
        {
            byte[] bMsg = DataEncoding(strMessage); 
            
            bMsg = esCapeChk(bMsg,2,2); // Escape
           
            for (int i = 0; i < bMsg.GetLength(0); i++)
            {
                // 1. 첫데이터가 STX 인지 확인 플래그 체크
                if (chkFlagData != true)
                {
                    // 1-1. 첫데이터가 STX?
                    if (bMsg[i] == STX)
                    {
                        // 1-1-1. STX가 연속으로 들어왔는지 확인
                        //        즉, 첫번째 데이터
                        if (read_msg_count >= 1)
                        {
                            readMsg[0] = bMsg[i];
                            read_msg_count = 1;
                            chkFlagStx = true;
                        }
                        // 1-1-2. STX가 연속으로 안들어왔를 경우
                        //        즉, 이상없이 데이터 저장
                        else
                        {
                            chkFlagData = true;

                            if (bMsg[i] == unUsed)
                            { chkFlagNeed = false; }

                            readMsg[read_msg_count++] = bMsg[i];
                        }
                    }
                    // 1-2. STX가 아닐경우
                    else
                    {
                        // 1-2-1. STX 체크
                        if (chkFlagStx == true)
                        {
                            // 쓰레기 데이터 버림
                        }
                        // 1-2-2. STX가 아닐 경우
                        else
                        {
                            if (bMsg[i] == unUsed)
                            { chkFlagNeed = false; }
                            chkFlagData = true;
                            readMsg[read_msg_count++] = bMsg[i];
                        }
                    }
                }
                // 2. ETX 값이 올때까지 저장
                else
                {
                    // 2-1. ETX 가 들어왔는지?
                    if (bMsg[i] == ETX)
                    {
                        readMsg[read_msg_count++] = bMsg[i];                      

                        if (chkFlagNeed == true )
                        {
                            RFDistribute(); // 전문분배
                        }
                        chkFlagNeed = true;

                        // 값 초기화
                        chkFlagStx = false;
                        chkFlagData = false;
                        read_msg_count = 0;
                    }
                    else
                    {
                        readMsg[read_msg_count++] = bMsg[i];
                    }
                }
            }
        }
        #endregion

        #region esCapeChk 메서드
        /// <summary>
        /// escape check
        /// </summary>
        /// <param name="bMsg"></param>
        /// <param name="Hpos"></param>
        /// <param name="Tpos"></param>
        /// <returns></returns>
        private byte[] esCapeChk(byte[] bMsg, int Hpos, int Tpos)
        {
	        byte[] bEscape = new byte[] { 0x7e, 0x7d }; // 삭제하고 싶은 Escape 문자를 추가
	        ArrayList obj = new ArrayList(); // array List Create

	        for (int i = 0; i < bMsg.Length; i++) // 수신받은 byte 전체를 arrylist에 저장
		        obj.Add(bMsg[i]);

            for (int i = Hpos; i < obj.Count - Tpos; i++)
	        {
		        for (int j = 0; j < bEscape.Length; j++)
		        {
			        if (byte.Parse(obj[i].ToString()) == bEscape[j])
			        {
				        obj.RemoveAt(i);
				        i = 0;
			        }
		        }
	        }

	        // escape를 제외한 개수로 배열을 새롭게 잡는다.
	        byte[] newMsg = new byte[obj.Count];

	        for (int i = 0; i < obj.Count; i++)
		        newMsg[i] = (byte)(obj[i]);

	        return newMsg;
        }
        #endregion

        #region 전문분배 메서드
        /// <summary>
        /// 전문분배
        /// </summary>
        private void RFDistribute()
        {
            string strBuffer = string.Empty;
            //전문 생성
            byte[] bbMsg = new byte[read_msg_count];
            for (int j = 0; j < read_msg_count; j++)
            {
                bbMsg[j] = readMsg[j];
                // 2009.10.26 KYY Add Test
                strBuffer += String.Format("{0:X2} ", bbMsg[j]);
            }           
            Data_Log_Error(strBuffer); // 2009.10.26 KYY Add Test
            strBuffer = string.Empty; // 2009.10.26 KYY Add Test

            //node_id, msg_Data
            // node_id : 2->온도,습도,조도 | 100->인체감지
            node_id = string.Format("{0:X2} ", readMsg[12]).Trim();
            // sensor_type : 1->습도, 2->온도, 3->조도
            sensor_type = string.Format("{0:X2} ", readMsg[16]).Trim();
            msg_bufferData1 = string.Format("{0:X2} ", readMsg[18]).Trim();
            msg_bufferData2 = string.Format("{0:X2} ", readMsg[19]).Trim();
            datavalue1 = msg_bufferData2 + msg_bufferData1;
            DateTime now_time = DateTime.Now;
            datatime = ((now_time)).ToLongTimeString();

            // 인체감지 이벤트를 감지한다.
            if (HexToDecimal(node_id) == 100) // 인체감지
            {
                bscreen = true;
            }
            else
            {
                bscreen = false;
                data_count++;
            }

            if (bscreen == true)
            {
                datavalue = HexToDecimal(datavalue1);

                // 파일 저장을 위해
                for (int k = 0; k < 10; k++)
                {
                    msg_bufferData3 = string.Format("{0:X2} ", readMsg[k * 2 + 19]).Trim() + string.Format("{0:X2} ", readMsg[k * 2 + 20]).Trim();
                    datavalue3 = HexToDecimal(msg_bufferData3);
                    arr_FileSave[k] = datavalue3.ToString();
                }

                SetMonitor(0); // 모니터 켜짐
                data_count++;
            }
            else
            {
                if (data_count > iscreensaver)
                {
                    SetMonitor(1); // 모니터 꺼짐
                    data_count = 0;
                }
            }
            
            switch (HexToDecimal(sensor_type)) // 온도,조도,습도
            {
                case 1: //습도
                    datavalue = HexToDecimal(datavalue1);

                    // 파일 저장을 위해
                    for (int k = 0; k < 10; k++)
                    {
                        msg_bufferData3 = string.Format("{0:X2} ", readMsg[k * 2 + 19]).Trim() + string.Format("{0:X2} ", readMsg[k * 2 + 20]).Trim();
                        datavalue3 = HexToDecimal(msg_bufferData3);
                        arr_FileSave[k] = humicalc(datavalue3);
                    }

                    SetText(humicalc(datavalue));
                    break;
                case 2: //온도
                    datavalue = HexToDecimal(datavalue1);

                    // 파일 저장을 위해
                    for (int k = 0; k < 10; k++)
                    {
                        msg_bufferData3 = string.Format("{0:X2} ", readMsg[k * 2 + 19]).Trim() + string.Format("{0:X2} ", readMsg[k * 2 + 20]).Trim();
                        datavalue3 = HexToDecimal(msg_bufferData3);
                        arr_FileSave[k] = tempcalc(datavalue3);
                    }

                    SetText(tempcalc(datavalue)); // 온도 값 출력
                    break;
                case 3: // 조도
                    // 조도값을 출력하기 위해선 이곳에 추가하시오
                    break;
                default: // 쓰레기
                    break;
            }
        }
        #endregion

        #region 모니터 절전기능 메서드
        /// <summary>
        /// 모니터 절전기능 함수
        /// </summary>
        /// <param name="node"></param>
        private void SetMonitor(int node)
        {
            string DataNow = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString();
            string PIRFilePath = Application.StartupPath + @"\Log\" + DataNow + "_PIR.txt";
            string PIRWriteLname = "     인체센서 :  ";

            if (this.InvokeRequired) 
            {
                SetMonitorCallback d = new SetMonitorCallback(SetMonitor);
                this.Invoke(d, new object[] { node });
            }
            else
            {
                // 로그 파일을 만들고 데이터를 저장한다.
                FileCreate(PIRFilePath, PIRWriteLname);
                // 로그 삭제
                FileDel(PIRFilePath);

                if (node == 1) // 모니터가 꺼질경우
                {
                    SendMessage(this.Handle.ToInt32(), WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_OFF);
                }
                else // 모니터가 켜질경우
                {
                    SendMessage(this.Handle.ToInt32(), WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_ON);
                }
            }
        }
        #endregion

        #region 로그관리 메서드
        /// <summary>
        /// 로그 관리
        /// </summary>
        private void Data_Log_Error(string errorMessage)
        {
            string DataNow = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString();
            string LogFile = Application.StartupPath + @"\Error\" + DataNow + "_errorLog.txt";

            FileStream fs = new FileStream(LogFile, FileMode.Append, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.UTF8);
            sw.WriteLine(DateTime.Now + "  : " + errorMessage);

            sw.Flush();
            sw.Close();
            fs.Close();

            FileDel(LogFile);
        }
        #endregion

        #region 파일 생성 및 출력 메서드

        /// <summary>
        /// 파일 생성 및 출력
        /// </summary>
        /// <param name="sFilepath"></param>
        /// <param name="WriteLname"></param>
        private void FileCreate(string sFilepath, string WriteLname)
        {
            try
            {
                FileStream fs = new FileStream(sFilepath, FileMode.Append, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.UTF8);
                sw.WriteLine(DateTime.Now + WriteLname + arr_FileSave[0] + "|" + arr_FileSave[1] + "|" + arr_FileSave[2] + "|" + arr_FileSave[3] + "|" + arr_FileSave[4] + "|" + arr_FileSave[5] + "|" + arr_FileSave[6] + "|" + arr_FileSave[7] + "|" + arr_FileSave[8] + "|" + arr_FileSave[9]);

                sw.Flush();
                sw.Close();
                fs.Close();
            }
            catch (System.Exception e)
            {
                Data_Log_Error(e.Message);
            }
        }
        #endregion

        #region 파일 삭제 메서드

        /// <summary>
        /// 파일 삭제
        /// </summary>
        /// <param name="sFilepath"></param>
        private void FileDel(string sFilepath)
        {
            try
            {
                if (1 == DateTime.Compare(DateTime.Now, File.GetCreationTime(sFilepath).AddDays(3)))
                {
                    File.Delete(sFilepath);
                }
            }
            catch (System.Exception e)
            {
                Data_Log_Error(e.Message);
            }
        }
        #endregion

        #region 온도, 습도 값 표시 메서드

        /// <summary>
        /// 온도, 습도 값 표시
        /// </summary>
        /// <param name="strLine"></param>
        private void SetText(string strLine)
        {
            string DataNow = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString();
            string HUMIFilePath = Application.StartupPath + @"\Log\" + DataNow + "_HUMI.txt";
            string HUMIWriteLname = "     습도 :  ";

            string TEMPFilePath = Application.StartupPath + @"\Log\" + DataNow + "_TEMP.txt";
            string TEMPWriteLname = "     온도 :  ";
            
            if (HexToDecimal(sensor_type) == 1)
            {
                if (this.label5.InvokeRequired) // 습도
                {
                    SetTextCallback d = new SetTextCallback(SetText);
                    this.label5.Invoke(d, new object[] { strLine });
                }
                else
                {
                    label5.Text = strLine + "%";

                    // 로그 파일을 만들고 데이터를 저장한다.
                    FileCreate(HUMIFilePath, HUMIWriteLname);
                    // 로그 삭제
                    FileDel(HUMIFilePath); 
                }
            }

            if (HexToDecimal(sensor_type) == 2)
            {
                if (this.label3.InvokeRequired) // 온도
                {
                    SetTextCallback d = new SetTextCallback(SetText);
                    this.Invoke(d, new object[] { strLine });
                }
                else
                {
                    label3.Text = strLine + "℃";

                    // 로그 파일을 만들고 데이터를 저장한다.
                    FileCreate(TEMPFilePath, TEMPWriteLname);
                    // 로그 삭제
                    FileDel(TEMPFilePath); 
                }
            }
        }
        #endregion

        #region 헥사 -> 데시멀 치환 메서드

        /// <summary>
        /// Hexa -> Decimal 치환
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        private int HexToDecimal(string str)
        {
            return Convert.ToInt32(str, 16);
        }
        #endregion

        private void Form1_Load(object sender, EventArgs e)
        {
            string LogFolder = Application.StartupPath + @"\Log";
            string ErrorFolder = Application.StartupPath + @"\Error";

            Directory.CreateDirectory(LogFolder);   // 로그 폴더 생성
            Directory.CreateDirectory(ErrorFolder); // 에러 폴더 생성

            Load_Setting();
            try
            {
                this.RS232.PortName = portselect;
                this.RS232.BaudRate = int.Parse(BautRateselect);
            }
            catch 
            {
                this.RS232.PortName = portselect;
                this.RS232.BaudRate = int.Parse(BautRateselect);
            }

            RS232_Open();
           
            Screen_Setting(int.Parse(MonitorMode));
            
            try
            {
                this.BackgroundImage = Image.FromFile(Application.StartupPath + Backimg);
            }
            catch
            {
                MessageBox.Show("실행파일이 있는 곳으로 그림파일을 옴기세요. 파일형식은 jpg");
            }

            //해상도에 따른 위치 변화
            this.label1.Location = new System.Drawing.Point(screenright - 887, 30); // 온도 레이블
            this.label3.Location = new System.Drawing.Point(screenright - 670, 33);  // 온도 값 670
            this.label2.Location = new System.Drawing.Point(screenright - 437, 30);  // 습도 레이블 480
            this.label5.Location = new System.Drawing.Point(screenright - 230, 33);  // 값 263
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            RS232_Close();
        }

        #region conf.xml 설정 메서드
        /// <summary>
        /// conf.xml 설정
        /// </summary>
        private void Load_Setting()
        {
            string FileName = Application.StartupPath + @"\conf.xml";

            try // xml 파일이 있을경우
            {
                XmlReader reader = XmlReader.Create(FileName);
                reader.Read();
                reader.ReadStartElement("Setup");

                reader.ReadStartElement("Display"); // 주/부 모니터 설정
                reader.ReadStartElement("mode");
                MonitorMode = reader.ReadString(); // primary:1, Second:0
                reader.ReadEndElement(); 
                reader.ReadEndElement(); 

                reader.ReadStartElement("Background"); // 폼 배경 스킨 설정
                reader.ReadStartElement("jpg");
                Backimg += reader.ReadString(); 
                reader.ReadEndElement(); 
                reader.ReadEndElement(); 

                reader.ReadStartElement("Port"); // 포트 설정
                reader.ReadStartElement("com");  
                portselect = reader.ReadString();
                reader.ReadEndElement();

                reader.ReadStartElement("BauteRate"); 
                BautRateselect = reader.ReadString();             

                reader.ReadEndElement();  
                reader.ReadEndElement(); 

                reader.ReadStartElement("Screen"); // 절전 딜레이 설정
                reader.ReadStartElement("time");
                screensaver = reader.ReadString();
                iscreensaver = int.Parse(screensaver);

                reader.ReadEndElement(); // jpg를 닫음
            }
            catch // xml 파일이 없을경우
            {
                MessageBox.Show("실행파일 있는곳으로" + FileName + "옴기세요");
            }
        }
        #endregion

        #region 습도 계산 메서드

        /// <summary>
        /// 습도계산
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private string humicalc(double value) // 습도계산 함수
        {
            cel = -4 + 0.0405 * (double)value + (-0.0000028 * (double)value) * (double)value;
            cel1 = (double_cel - 25) * (0.01 + 0.00008 * (double)value) + cel;
            //string celbuffer1 = string.Format("{0:##.00}", cel1);
            string celbuffer1 = string.Format("{0:##}", cel1);
            return celbuffer1;
        }
        #endregion

        #region 온도 계산 메서드

        /// <summary>
        /// 온도계산
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private string tempcalc(double value) // 온도계산 함수
        {
            cel = -39.60 + 0.01 * (double)value;
            double_cel = cel;
            //string celbuffer = string.Format("{0:##.00}", cel);
            string celbuffer = string.Format("{0:##}", cel);
            return celbuffer;
        }
        #endregion

        #region 해상도 자동조절 메서드

        /// <summary>
        /// 해상도 자동조절 메서드
        /// </summary>
        /// <param name="number"></param>
        private void Screen_Setting(int number)
        {
            // 시스템의 모든 디스플레이 배열을 가져옵니다.
            Screen[] screens = Screen.AllScreens;
            int upperBound = screens.GetUpperBound(0);

            // 크기 구하기
            this.Width = screens[number].Bounds.Width;
            this.Height = screens[number].Bounds.Height;

            // 위치
            this.Left = screens[number].Bounds.Left;
            this.Top = screens[number].Bounds.Top; // 위는 고정

            screenright  =  screens[number].Bounds.Width;
            screentop = screens[number].Bounds.Height;
        }
        #endregion

        private void Form1_DoubleClick(object sender, EventArgs e)
        {
            try // 인덱스 배열 초과 현상 때문에... 2009.10.13 KYY 추가
            {
                if (MessageBox.Show("종료할까요?", "종료", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
                {
                    Application.Exit();
                }
            }
            catch (System.Exception l)
            {
                Data_Log_Error(l.Message);
            }
        }

        #region 사용하지 않는 이벤트

        // 폼 마우스 이동시 주석을 풀면 됩니다.
        //private void Form1_MouseMove(object sender, MouseEventArgs e)
        //{
        //    if (isMouseDown)
        //    {
        //        // Set the form's location property to the new position.
        //        Point mousePos = Control.MousePosition;
        //        mousePos.Offset(mouseOffset.X, mouseOffset.Y);
        //        this.Location = mousePos;
        //    }
        //}

        //private void Form1_MouseUp(object sender, MouseEventArgs e)
        //{
        //    int xOffset;
        //    int yOffset;
        //    if (e.Button == MouseButtons.Left)
        //    {
        //        // Assign coordinates to mouseOffset variable based on 
        //        // current position of the mouse pointer.
        //        xOffset = -e.X - SystemInformation.FrameBorderSize.Width + 5;
        //        //yOffset = e.Y - SystemInformation.FrameBorderSize.Height
        //        //    - SystemInformation.CaptionButtonSize.Height;
        //        yOffset = -e.Y - SystemInformation.CaptionHeight -
        //            SystemInformation.FrameBorderSize.Height + 30;
        //        mouseOffset = new Point(xOffset, yOffset);
        //        isMouseDown = true;
        //    }  
        //}

        //private void Form1_MouseDown(object sender, MouseEventArgs e)
        //{
        //    if (e.Button == MouseButtons.Left)
        //    {
        //        isMouseDown = false;
        //    }
        //}
#endregion
    }
}
