using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO.Ports;
using System.IO;
using System.Threading;
using System.Windows.Threading;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using System.Management;
using System.Windows.Interop;
using System.Collections;

namespace VSP
{

    /// <summary>
    /// SerialBasic.xaml 的交互逻辑
    /// </summary>
    public partial class SerialBasic : UserControl
    {
        #region 变量定义

        #region 内部变量
        private bool Listening = false;
        //是否没有执行完invoke相关操作   
        private bool Closing = false;
        private SerialPort serial = null;
        private StringBuilder builder = new StringBuilder();
        private bool hexadecimalDisplay = false;
        private DispatcherTimer autoSendTimer = new DispatcherTimer();
        private DispatcherTimer autoDetectionTimer = new DispatcherTimer();
        private Queue sendQueue = new Queue();
        private Thread threadSendQueue;
        static UInt32 receiveBytesCount = 0;
        static UInt32 sendBytesCount = 0;
        //接收数据
        private delegate void UpdateUiTextDelegate(string text);
        #endregion

        #endregion

        public SerialBasic()
        {
            InitializeComponent();
            serial = new SerialPort();
            serial.Encoding = Encoding.Default;
            serial.DataReceived += ReceiveData;
            serial.ReadBufferSize = 4096;
            serial.ReceivedBytesThreshold = 1;
            serial.ReadTimeout = 20;
            serial.NewLine = "\n";
          
            readUserData();
            receiveColorCheckBox_Config();

            search_serial();  //通过WMI获取COM端口 
            threadSendQueue = new Thread(SendQueueTask);
            threadSendQueue.Start();

            //添加设备变化通知
            MainWindowMonitor.DeviceChangedEvent += SerialUpdate;
            MainWindowMonitor.WindowClosedEvent += saveUserData;
            //设置状态栏提示
            statusTextBlock.Text = "准备就绪";
        }


        private void receiveColorCheckBox_Config()
        {
            if (receiveColorCheckBox.IsChecked == true)
            {
                receiveTextBox.Background = new SolidColorBrush(Colors.White);
                receiveTextBox.Foreground = new SolidColorBrush(Colors.Black);
            }
            else
            {
                receiveTextBox.Background = new SolidColorBrush(Colors.Black);
                receiveTextBox.Foreground = new SolidColorBrush(Colors.White);
            }
        }

        private void readUserData()
        {
            //if (portNamesCombobox.Items.Contains(Properties.Settings.Default.serialName) == true)
            //{
            //    portNamesCombobox.SelectedIndex = portNamesCombobox.Items.IndexOf(Properties.Settings.Default.serialName);
            //}
            portNamesCombobox.Items.Add(Properties.Settings.Default.serialName);
            autoSendCycleTextBox.Text = Properties.Settings.Default.autoSendCycle;
            sendTextBox.Text = Properties.Settings.Default.sendString;
            baudRateCombobox.Text = Properties.Settings.Default.BaudRate;
            showSendCheckBox.IsChecked = Properties.Settings.Default.showSendCheckBox;
            sendNewLineCheckBox.IsChecked = Properties.Settings.Default.sendNewLineCheckBox;
            receiveColorCheckBox.IsChecked = Properties.Settings.Default.receiveColorCheckBox;

        }

        private void saveUserData(object sender, EventArgs e)
        {
            Properties.Settings.Default.autoSendCycle = autoSendCycleTextBox.Text;
            Properties.Settings.Default.serialName = portNamesCombobox.Text;
            Properties.Settings.Default.sendString = sendTextBox.Text;
            Properties.Settings.Default.BaudRate = baudRateCombobox.Text;
            Properties.Settings.Default.showSendCheckBox = (bool)showSendCheckBox.IsChecked;
            Properties.Settings.Default.sendNewLineCheckBox = (bool)sendNewLineCheckBox.IsChecked;
            Properties.Settings.Default.receiveColorCheckBox = (bool)receiveColorCheckBox.IsChecked;

            Properties.Settings.Default.Save();

        }


        #region 自动更新串口号

        //搜索串口名
        private void search_serial()
        {
            Thread threadDeveiceChanged = new Thread(() =>
            {
                AddValuablePortName(MainWindowMonitor.MulGetHardwareInfo(MainWindowMonitor.HardwareEnum.Win32_PnPEntity, "Name"));
            });
            threadDeveiceChanged.Start();
        }

        private void SerialUpdate(object sender, EventArgs e)
        {
            Thread threadDeveiceChanged = new Thread(() =>
            {
                string[] portnames = MainWindowMonitor.MulGetHardwareInfo(MainWindowMonitor.HardwareEnum.Win32_PnPEntity, "Name");
                Dispatcher.Invoke(new Action(() =>
                {
                    if (turnOnButton.Content.Equals("关闭串口"))
                    {
                        bool flag = true;
                        //查找所有存在的串口                  

                        for (int i = 0; i < portnames.Length; i++)
                        {
                            if (portnames[i].Contains(serial.PortName))
                            {
                                flag = false;//不是本串口被拔      
                            }
                        }
                        if (flag == true)//所有存在的串口中找不到已经打开的串口      
                        {
                            Close_Port();
                            AddValuablePortName(portnames);
                            statusTextBlock.Text = "设备被拔出！";
                        }
                    }
                    else AddValuablePortName(portnames);
                }));
            });
            threadDeveiceChanged.Start();
        }

        private void AddValuablePortName(string[] serialPortName)
        {
            Dispatcher.Invoke(new Action(delegate
            {
                string Com_Now = "";
                if (portNamesCombobox.Items.Count > 0)
                {
                    Com_Now = portNamesCombobox.Text;
                }
                portNamesCombobox.Items.Clear();
                foreach (string name in serialPortName)
                {
                    portNamesCombobox.Items.Add(name);
                }
                if (portNamesCombobox.Items.Count == 0)
                {
                    portNamesCombobox.Items.Add("");
                }

                if (portNamesCombobox.Items.IndexOf(Com_Now) >= 0)
                {
                    portNamesCombobox.SelectedIndex = portNamesCombobox.Items.IndexOf(Com_Now);

                }
                else
                {
                    portNamesCombobox.SelectedIndex = portNamesCombobox.Items.Count > 0 ? 0 : -1;
                }

            }));
        }



        #endregion

        #region 串口配置面板

        //使能或关闭串口配置相关的控件
        private void serialSettingControlState(bool state)
        {
            portNamesCombobox.IsEnabled = state;
            baudRateCombobox.IsEnabled = state;
            parityCombobox.IsEnabled = state;
            dataBitsCombobox.IsEnabled = state;
            stopBitsCombobox.IsEnabled = state;
        }

        private bool Open_Port()
        {
            try
            {
                sendQueue.Clear();
                //配置串口
                serial.PortName = portNamesCombobox.Text.Substring(0, portNamesCombobox.Text.IndexOf(":"));
                serial.BaudRate = Convert.ToInt32(baudRateCombobox.Text);
                serial.Parity = (Parity)Enum.Parse(typeof(Parity), parityCombobox.Text);
                serial.DataBits = Convert.ToInt16(dataBitsCombobox.Text);
                serial.StopBits = (StopBits)Enum.Parse(typeof(StopBits), stopBitsCombobox.Text);

                //开启串口
                serial.Open();

                //关闭串口配置面板
                serialSettingControlState(false);

                statusTextBlock.Text = "串口已开启";

                //显示提示文字
                turnOnButton.Content = "关闭串口";

                serialPortStatusEllipse.Fill = Brushes.Red;

                //使能发送面板
                // sendControlBorder.IsEnabled = true;
                return true;

            }
            catch (Exception ex)
            {
                ex.ToString();
                statusTextBlock.Text = ex.Message;
                //statusTextBlock.Text = "串口打开失败";
                return false;
            }
        }
        private void Close_Port()
        {
            try
            {
                //关闭定时器
                autoSendTimer.Stop();
                autoSendCheckBox.IsChecked = false;
                sendQueue.Clear();
                //serial.DiscardOutBuffer();//清发送缓存
                //serial.DiscardInBuffer();//清接收缓存
                Closing = true;
                while (Listening) { Application.Current.Dispatcher.Invoke(DispatcherPriority.Background, new ThreadStart(delegate { })); };
                serial.Close();              
                Closing = false;
                //使能串口配置面板
                serialSettingControlState(true);

                statusTextBlock.Text = "串口已关闭";

                //显示提示文字
                turnOnButton.Content = "打开串口";

                serialPortStatusEllipse.Fill = Brushes.Gray;
                //使能发送面板
                //sendControlBorder.IsEnabled = false;
            }
            catch (Exception ex)
            {
                statusTextBlock.Text = ex.Message;
            }

        }
        private void SerialClick(object sender, RoutedEventArgs e)
        {
            if (turnOnButton.Content.Equals("打开串口"))
                Open_Port();
            else Close_Port();

        }
        private void TurnOnButton_Checked(object sender, RoutedEventArgs e)
        {

        }


        //关闭串口
        private void TurnOnButton_Unchecked(object sender, RoutedEventArgs e)
        {

        }

        #endregion

        #region 接收显示窗口

        private void ReceiveData(object sender, SerialDataReceivedEventArgs e)
        {
            if (Closing) return;
            Listening = true;
            if (hexadecimalDisplay == true)
            {
                int n = serial.BytesToRead;//先记录下来，避免某种原因，人为的原因，操作几次之间时间长，缓存不一致 
                byte[] buf = new byte[n];//声明一个临时数组存储当前来的串口数据               
                serial.Read(buf, 0, n);//读取缓冲数据
                receiveBytesCount += (UInt32)n;//增加接收计数                               
                builder.Clear();//清除字符串构造器的内容
                //依次的拼接出16进制字符串  
                foreach (byte b in buf)
                {
                    builder.Append(b.ToString("X2") + " ");
                }
                string hexString = builder.ToString();
                Dispatcher.Invoke(() =>
                {
                    if (stopShowingButton.IsChecked == false) //没有关闭数据显示
                    {
                        receiveTextBox.AppendText(hexString);
                    }
                    statusReceiveByteTextBlock.Text = receiveBytesCount.ToString();
                });

            }
            else
            {
                bool isShowTime = false;
                //Dispatcher.Invoke(() =>
                //{
                //    isShowTime = (bool)showTimeCheckBox.IsChecked;
                //});
                string receiveData = serial.ReadExisting();
                //try
                //{
                //    if (isShowTime) receiveData = DateTime.Now.ToLocalTime().ToString() + " : " +serial.ReadLine() + "\n";
                //    else receiveData = serial.ReadExisting();
                //}
                //catch { };
                //更新接收字节数
                receiveBytesCount += (UInt32)receiveData.Length;
                Dispatcher.Invoke(() =>
                {
                    //没有关闭数据显示
                    if (stopShowingButton.IsChecked == false)
                    {
                        receiveTextBox.AppendText(receiveData);
                    }
                    statusReceiveByteTextBlock.Text = receiveBytesCount.ToString();
                });
            }
            Listening = false;
        }
        public bool IsVerticalScrollBarAtButtom
        {
            get
            {
                bool isAtButtom = false;

                // get the vertical scroll position
                double dVer = receiveTextBox.VerticalOffset;

                //get the vertical size of the scrollable content area
                double dViewport = receiveTextBox.ViewportHeight;

                //get the vertical size of the visible content area
                double dExtent = receiveTextBox.ExtentHeight;

                if (dVer != 0)
                {
                    if (dVer + dViewport >= dExtent - 50)
                    {
                        isAtButtom = true;
                    }
                    else
                    {
                        isAtButtom = false;
                    }
                }
                else
                {
                    isAtButtom = false;
                }

                if (dViewport >= dExtent) isAtButtom = true;

                if (receiveTextBox.VerticalScrollBarVisibility == ScrollBarVisibility.Disabled
                    || receiveTextBox.VerticalScrollBarVisibility == ScrollBarVisibility.Hidden)
                {
                    isAtButtom = true;
                }


                return isAtButtom;
            }
        }

        //设置滚动条显示到末尾
        private void ReceiveTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            const int MaxLength = 50000;
            try
            {
                if (IsVerticalScrollBarAtButtom == true) receiveTextBox.ScrollToEnd();
                //if (receiveTextBox.Text.Length > MaxLength)
                //{
                //    receiveTextBox.Text = receiveTextBox.Text.Substring(MaxLength / 10, receiveTextBox.Text.Length - MaxLength / 10);
                //}
            }
            catch
            {             
            }
        }


        #endregion

        #region 接收设置面板

        //清空接收数据
        private void ClearReceiveButton_Click(object sender, RoutedEventArgs e)
        {
            receiveTextBox.Clear();
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            sendTextBox.Clear();
        }

        #endregion

        #region 发送控制面板

        private struct SendSetStr//发送数据线程传递参数的结构体格式
        {
            public string SendSetData;//发送的数据
            public bool SendAscill;//发送模式
        }
        private void SendQueueTask()
        {
            while (true)
            {
                if (sendQueue.Count <= 0)
                {
                    Thread.Sleep(100);
                    continue;
                }
                
                SendSetStr SendSet = (SendSetStr)sendQueue.Dequeue();
                string sendData = SendSet.SendSetData;
                if (SendSet.SendAscill)
                {
                    try
                    {
                        serial.Write(sendData);
                        Dispatcher.Invoke(() =>
                        {
                            if (showSendCheckBox.IsChecked == true)
                            {
                                receiveTextBox.AppendText(sendData);
                            }
                            //更新发送数据计数
                            sendBytesCount += (UInt32)sendData.Length;
                            statusSendByteTextBlock.Text = sendBytesCount.ToString();
                            statusTextBlock.Text = "发送成功";
                        });
                    }
                    catch (Exception ex)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            autoSendCheckBox.IsChecked = false;//关闭自动发送
                            statusTextBlock.Text = ex.Message;
                        });
                    }
                }
                else
                {
                    try
                    {
                        sendData = sendData.Replace(" ", "");
                        sendData = sendData.Replace("0X", "");
                        sendData = sendData.Replace("0x", "");
                        sendData = sendData.Replace("\r", "");
                        sendData = sendData.Replace("\n", "");
                        sendData = sendData.Replace("\t", "");

                        byte[] sendByte = new byte[sendData.Length / 2 + sendData.Length % 2];
                        for (int i = 0; i < (sendData.Length - sendData.Length % 2) / 2; i++)//转换偶数个
                        {
                            sendByte[i] = Convert.ToByte(sendData.Substring(i * 2, 2), 16);    //转换               
                        }
                        if (sendData.Length % 2 != 0)//单独处理最后一个字符    
                        {
                            sendByte[sendByte.Length - 1] = Convert.ToByte(sendData.Substring(sendData.Length - 1, 1), 16);
                        }

                        try
                        {
                            serial.Write(sendByte, 0, sendByte.Length);
                            //更新发送数据计数
                            sendBytesCount += (UInt32)sendByte.Length;
                            Dispatcher.Invoke(() =>
                            {
                                statusSendByteTextBlock.Text = sendBytesCount.ToString();
                                if (showSendCheckBox.IsChecked == true) receiveTextBox.AppendText("\n" + sendData + "\n");
                                statusTextBlock.Text = "发送成功";
                            });
                        }
                        catch (Exception ex)
                        {
                            Dispatcher.Invoke(() =>
                            {
                                statusTextBlock.Text = ex.Message;
                                autoSendCheckBox.IsChecked = false;//关闭自动发送
                            });
                        }
                    }
                    catch (Exception e)
                    {
                        e.ToString();
                        Dispatcher.Invoke(() =>
                        {
                            autoSendCheckBox.IsChecked = false;//关闭自动发送
                            statusTextBlock.Text = "请输入16进制";
                        });
                    }
                }

            }
        }
        private void SerialPortSendAsString(string sendData)
        {
            if (serial.IsOpen == false)
            {
                //statusTextBlock.Text = "请先打开串口";
                //return;
                if (Open_Port() == false) return;
            }
            if (sendQueue.Count > 10)
            {
                statusTextBlock.Text = "发送队列已满";
                return;
            }
            SendSetStr SendSet = new SendSetStr();
            SendSet.SendSetData = sendData;
            SendSet.SendAscill = true;
            sendQueue.Enqueue(SendSet);
            return;
            try
            {
                serial.Write(sendData);
                if (showSendCheckBox.IsChecked == true)
                {
                    receiveTextBox.AppendText(sendData);
                }
                //更新发送数据计数
                sendBytesCount += (UInt32)sendData.Length;
                statusSendByteTextBlock.Text = sendBytesCount.ToString();
                statusTextBlock.Text = "发送成功";
            }
            catch (Exception ex)
            {
                autoSendCheckBox.IsChecked = false;//关闭自动发送
                statusTextBlock.Text = ex.Message;
            }
        }
        //发送数据
        private void SerialPortSendAsHex(string sendData)
        {
            if (serial.IsOpen == false)
            {
                //statusTextBlock.Text = "请先打开串口";
                //return;
                if (Open_Port() == false) return;
            }
            if (sendQueue.Count > 10)
            {
                statusTextBlock.Text = "发送队列已满";
                return;
            }
            SendSetStr SendSet = new SendSetStr();
            SendSet.SendSetData = sendData;
            SendSet.SendAscill = false;
            sendQueue.Enqueue(SendSet);
            return;
            try
            {
                sendData = sendData.Replace(" ", "");
                sendData = sendData.Replace("0X", "");
                sendData = sendData.Replace("0x", "");
                sendData = sendData.Replace("\r", "");
                sendData = sendData.Replace("\n", "");
                sendData = sendData.Replace("\t", "");

                byte[] sendByte = new byte[sendData.Length / 2 + sendData.Length % 2];
                for (int i = 0; i < (sendData.Length - sendData.Length % 2) / 2; i++)//转换偶数个
                {
                    sendByte[i] = Convert.ToByte(sendData.Substring(i * 2, 2), 16);    //转换               
                }
                if (sendData.Length % 2 != 0)//单独处理最后一个字符    
                {
                    sendByte[sendByte.Length - 1] = Convert.ToByte(sendData.Substring(sendData.Length - 1, 1), 16);
                }
                try
                {
                    serial.Write(sendByte, 0, sendByte.Length);
                    //更新发送数据计数
                    sendBytesCount += (UInt32)sendByte.Length;
                    statusSendByteTextBlock.Text = sendBytesCount.ToString();
                    if (showSendCheckBox.IsChecked == true) receiveTextBox.AppendText("\n" + sendData + "\n");
                }
                catch (Exception ex)
                {
                    statusTextBlock.Text = ex.Message;
                    autoSendCheckBox.IsChecked = false;//关闭自动发送
                    return;
                }
                statusTextBlock.Text = "发送成功";

            }
            catch (Exception e)
            {
                e.ToString();
                autoSendCheckBox.IsChecked = false;//关闭自动发送
                statusTextBlock.Text = "请输入16进制";
                return;
            }
        }
        private void SerialPortSend(string sendData)
        {
            //字符串发送
            if (hexadecimalSendCheckBox.IsChecked == false)
            {
                if (sendNewLineCheckBox.IsChecked == true)
                {
                    sendData += "\r\n";
                }
                SerialPortSendAsString(sendData);
            }
            else
            {
                SerialPortSendAsHex(sendData);
            }
        }
        //手动发送数据
        private void SendButton_Click(object sender, RoutedEventArgs e)
        {
            if (sendTextBox.Text != string.Empty) SerialPortSend(sendTextBox.Text);
            else statusTextBlock.Text = "发送区为空";

        }
        void AutoSendTimer_Tick(object sender, EventArgs e)
        {
            if (sendTextBox.Text != string.Empty) SerialPortSend(sendTextBox.Text);
            else
            {
                Close_Port();
                statusTextBlock.Text = "发送区为空";
                return;
            }
            //设置新的定时时间     
            try
            {
                autoSendTimer.Interval = new TimeSpan(0, 0, 0, 0, Convert.ToInt32(autoSendCycleTextBox.Text));
            }
            catch (Exception ex)
            {
                Close_Port();
                statusTextBlock.Text = ex.Message;
                return;
            }

        }
        //设置自动发送定时器
        private void AutoSendCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            //创建定时器
            autoSendTimer.Tick += new EventHandler(AutoSendTimer_Tick);

            //设置定时时间，开启定时器
            try
            {
                autoSendTimer.Interval = new TimeSpan(0, 0, 0, 0, Convert.ToInt32(autoSendCycleTextBox.Text));
                autoSendTimer.Start();
            }
            catch (Exception ex)
            {
                Close_Port();
                statusTextBlock.Text = ex.Message;
                return;
            }
        }
        //自动发送时间到

        //关闭自动发送定时器
        private void AutoSendCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            autoSendTimer.Stop();
        }

        private void ClearSendButton_Click(object sender, RoutedEventArgs e)
        {
            sendTextBox.Clear();
        }
        #endregion


        private void SendTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {

             e.Handled = false;
            
        }

        private void FileOpen(object sender, ExecutedRoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.FileName = "serialCom";
            openFile.DefaultExt = ".txt";
            openFile.Filter = "TXT文本|*.txt";
            if (openFile.ShowDialog() == true)
            {
                sendTextBox.Text = File.ReadAllText(openFile.FileName, System.Text.Encoding.Default);

                fileNameTextBox.Text = openFile.FileName;
            }
        }

        private void FileSave(object sender, ExecutedRoutedEventArgs e)
        {

            if (receiveTextBox.Text == string.Empty)
            {
                statusTextBlock.Text = "接收区为空，不保存！";
            }
            else
            {
                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "TXT文本|*.txt";
                saveFile.FileName = "receiveText";
                if (saveFile.ShowDialog() == true)
                {
                    File.AppendAllText(saveFile.FileName, "\r\n-----------" + DateTime.Now.ToString() + "-----------\r\n");
                    File.AppendAllText(saveFile.FileName, receiveTextBox.Text);
                    statusTextBlock.Text = "保存成功！";
                }
            }
        }


        private void WindowClosed(object sender, ExecutedRoutedEventArgs e)
        {
            ;
        }

        private void SerialWindowUnloaded(object sender, RoutedEventArgs e)
        {
            ;
        }
        //清空计数
        private void countClearButton_Click(object sender, RoutedEventArgs e)
        {
            //接收、发送计数清零
            receiveBytesCount = 0;
            sendBytesCount = 0;

            //更新数据显示
            statusReceiveByteTextBlock.Text = receiveBytesCount.ToString();
            statusSendByteTextBlock.Text = sendBytesCount.ToString();

        }

        private void StopShowingButton_Checked(object sender, RoutedEventArgs e)
        {
            stopShowingButton.Content = "恢复显示";
        }

        private void StopShowingButton_Unchecked(object sender, RoutedEventArgs e)
        {
            stopShowingButton.Content = "停止显示";
        }
        #region 命令按钮
        private void Button_RST(object sender, RoutedEventArgs e)
        {
            SerialPortSendAsString("#r\r\n");
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            SerialPortSendAsString("#t" + DateTime.Now.ToLongTimeString().ToString().PadLeft(8, '0') + "\r\n");
            Thread.Sleep(10);
            SerialPortSendAsString("#y" + DateTime.Now.ToString("yy-MM-dd") + "\r\n");
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            SerialPortSendAsString("#c\r\n");
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            SerialPortSendAsString("#b\r\n");
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            SerialPortSendAsString("#s\r\n");
        }

        private void Button_Click_5(object sender, RoutedEventArgs e)
        {
            SerialPortSendAsString("#e\r\n");
        }

        private void Button_Click_6(object sender, RoutedEventArgs e)
        {
            SerialPortSendAsString("#w\r\n");
        }

        private void Button_Click_7(object sender, RoutedEventArgs e)
        {
            SerialPortSendAsString("#f\r\n");
        }

        private void Button_Click_8(object sender, RoutedEventArgs e)
        {
            SerialPortSendAsString("#p\r\n");
        }

        private void Button_Click_9(object sender, RoutedEventArgs e)
        {
            SerialPortSendAsString("#y\r\n");
        }

        private void Button_Click_10(object sender, RoutedEventArgs e)
        {
            SerialPortSendAsString("#n\r\n");
        }
        #endregion
        private void autoClearCheckBox_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void receiveColorCheckBox_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void receiveColorCheckBox_Click(object sender, RoutedEventArgs e)
        {
            receiveColorCheckBox_Config();
        }


        private void portNamesCombobox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                // serial.PortName = portNamesCombobox.Text.Substring(portNamesCombobox.Text.IndexOf("(") + 1, portNamesCombobox.Text.IndexOf(")") - portNamesCombobox.Text.IndexOf("(") - 1);
            }
            catch
            {

            }
        }

        private void baudRateCombobox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                // serial.BaudRate = Convert.ToInt32(baudRateCombobox.SelectedItem);
            }
            catch
            {
                ;
            }

        }

        private void parityCombobox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                // serial.Parity = (Parity)Enum.Parse(typeof(Parity), parityCombobox.Text);
            }
            catch
            {

            }
        }

        private void dataBitsCombobox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                //  serial.DataBits = Convert.ToInt16(dataBitsCombobox.Text);
            }
            catch
            {

            }
        }

        private void stopBitsCombobox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                //  serial.StopBits = (StopBits)Enum.Parse(typeof(StopBits), stopBitsCombobox.Text);
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
        }

        private void portNamesCombobox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            search_serial();
        }


        private void hexadecimalDisplayCheckBox_Click(object sender, RoutedEventArgs e)
        {
            hexadecimalDisplay = (bool)hexadecimalDisplayCheckBox.IsChecked;
        }
    }
}
