using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;


namespace VSP
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();
            this.Closing += Window_Closing;
        }

        private void SerialBasic_Loaded(object sender, RoutedEventArgs e)
        {
            HwndSource hwndSource = PresentationSource.FromVisual(this) as HwndSource;//窗口过程
            if (hwndSource != null)
            {
                hwndSource.AddHook(new HwndSourceHook(DeveiceChanged));  //挂钩
            }

        }
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (MessageBox.Show("是否要关闭？", "提示", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                e.Cancel = false;

            }
            else
            {
                e.Cancel = true;

            }
            MainWindowMonitor.WindowClosedEventFunction();

        }



        private IntPtr DeveiceChanged(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == MainWindowMonitor.WM_DEVICECHANGE)
            {             
                switch (wParam.ToInt32())
                {
                    case MainWindowMonitor.DBT_DEVICEARRIVAL://设备插入  
                        MainWindowMonitor.DeviceChangedFunction();


                        break;
                    case MainWindowMonitor.DBT_DEVICEREMOVECOMPLETE: //设备卸载
                        MainWindowMonitor.DeviceChangedFunction();

                        break;
                    default:
                        break;
                }
            }
            return IntPtr.Zero;
        }


    }
}
