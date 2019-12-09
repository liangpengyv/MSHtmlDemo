using System.IO;
using System.Text;
using System.Windows;
using mshtml;
using SHDocVw;

namespace MSHtmlDemo
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void On_Selector_Click(object sender, RoutedEventArgs e)
        {
            InternetExplorer ieWebBrowser = GetIEWindow();

            if (ieWebBrowser != null)
            {
                // 最小化窗口
                this.WindowState = WindowState.Minimized;

                // 读取欲注入的 JS 脚本
                StreamReader reader = new StreamReader("Selector.js", Encoding.UTF8);
                string strSelectorJS = reader.ReadToEnd();
                reader.Close();

                // 向 IE 注入 JS 脚本
                IHTMLWindow2 window2 = (ieWebBrowser.Document as IHTMLDocument2).parentWindow;
                window2.execScript(strSelectorJS);
            }
            else
            {
                MessageBox.Show("IE 浏览器未启动！");
            }
        }

        /// <summary>
        /// 获取 IE 进程窗口对象
        /// </summary>
        /// <returns></returns>
        private InternetExplorer GetIEWindow()
        {
            ShellWindowsClass shellWindows = new ShellWindowsClass();

            foreach (InternetExplorer ie in shellWindows)
            {
                if (ie.FullName.ToUpper().IndexOf("IEXPLORE.EXE") > 0)
                {
                    // 找到 IE 浏览器窗口
                    return ie;
                }
            }

            return null;
        }
    }
}
