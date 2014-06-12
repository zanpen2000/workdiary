using ClassLibrary;
using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.Configuration;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;

namespace WorkDiary
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window, INotifyPropertyChanged
    {
        public string MailUser { get; set; }
        public string MailTo { get; set; }
        public string LastFileName { get; set; }
        public string ContentCell { get; set; }

        RoutedCommand BrowseCommand = new RoutedCommand("Browse", typeof(MainWindow));
        RoutedCommand ReadCommand = new RoutedCommand("Read", typeof(MainWindow));
        RoutedCommand MailSendCommand = new RoutedCommand("SendMail", typeof(MainWindow));
        RoutedCommand ReadExcelCommand = new RoutedCommand("ReadExcel", typeof(MainWindow));


        private Person _person;
        public Person Person
        {
            get { return _person; }
            set
            {
                _person = value;
                if (PropertyChanged != null)
                {
                    PropertyChanged(this, new PropertyChangedEventArgs("Person"));
                }
            }
        }

        public MainWindow()
        {
            InitializeComponent();

            var task = Task.Run(() =>
            {
                this.Dispatcher.InvokeAsync((System.Action)delegate
                {
                    MailUser = ConfigurationManager.AppSettings.Get("mailuser");
                    MailTo = ConfigurationManager.AppSettings.Get("mailto");
                    LastFileName = ConfigurationManager.AppSettings.Get("lastfilename");
                    ContentCell = ConfigurationManager.AppSettings.Get("contentcell");

                    Binding b1 = new Binding("Text") { Source = this.oriExcelFile };
                    Binding b2 = new Binding("Text") { Source = this.personUI.tDate };
                    MultiBinding mb = new MultiBinding();
                    mb.Bindings.Add(b1);
                    mb.Bindings.Add(b2);

                    mb.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged;
                    mb.Converter = new FileDateConverter();
                    this.tNewFileName.SetBinding(System.Windows.Controls.TextBox.TextProperty, mb);

                    SetCommandBinding();
                });
            });

            task.GetAwaiter().OnCompleted(() =>
            {
                DataContext = this;
            });


        }

        private void SetCommandBinding()
        {
            this.btnBrowser.Command = BrowseCommand;
            this.BrowseCommand.InputGestures.Add(new KeyGesture(Key.B, ModifierKeys.Alt));
            this.btnBrowser.CommandTarget = this.oriExcelFile;

            CommandBinding browseCmdBinding = new CommandBinding();
            browseCmdBinding.Command = BrowseCommand;
            browseCmdBinding.Executed += cmdBinding_Executed;


            this.btnRead.Command = ReadCommand;
            this.ReadCommand.InputGestures.Add(new KeyGesture(Key.G, ModifierKeys.Alt));
            this.btnRead.CommandTarget = this.g1;
            CommandBinding readCmdBinding = new CommandBinding();
            readCmdBinding.Command = ReadCommand;
            readCmdBinding.Executed += readCmdBinding_Executed;
            readCmdBinding.CanExecute += readCmdBinding_CanExecute;

            this.btnSend.Command = MailSendCommand;
            MailSendCommand.InputGestures.Add(new KeyGesture(Key.M, ModifierKeys.Alt));
            this.btnSend.CommandTarget = this.tReceiver;
            CommandBinding mailSendCmdBinding = new CommandBinding();
            mailSendCmdBinding.Command = MailSendCommand;
            mailSendCmdBinding.CanExecute += mailSendCmdBinding_CanExecute;
            mailSendCmdBinding.Executed += mailSendCmdBinding_Executed;


            CommandBinding saveCommandBinding = new CommandBinding();
            saveCommandBinding.Command = ApplicationCommands.Save;
            btnSaveAs.Command = ApplicationCommands.Save;
            btnSaveAs.CommandTarget = this.tNewFileName;

            saveCommandBinding.CanExecute += saveCommandBinding_CanExecute;
            saveCommandBinding.Executed += saveCommandBinding_Executed;

            this.mainWindow.CommandBindings.Add(browseCmdBinding);
            this.mainWindow.CommandBindings.Add(readCmdBinding);
            this.mainWindow.CommandBindings.Add(mailSendCmdBinding);
            this.mainWindow.CommandBindings.Add(saveCommandBinding);

            this.oriExcelFile.TextChanged += (o, e) =>
            {
                if (System.IO.File.Exists(this.oriExcelFile.Text))
                {
                    ReadExcel(this.oriExcelFile.Text);
                }
            };
        }

        async void saveCommandBinding_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            await SaveAsExcel(tNewFileName.Text);
            await _saveConfig();
            e.Handled = true;
        }

        void saveCommandBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            bool result = false;

            if (!string.IsNullOrEmpty(oriExcelFile.Text) && !string.IsNullOrEmpty(tNewFileName.Text))
            {
                if (Person != null && !String.IsNullOrEmpty(Person.Date))
                {
                    result = DateTime.Parse(Person.Date).ToString("yyyy/MM/dd").Equals(DateTime.Now.ToString("yyyy/MM/dd"));
                }
            }
            e.CanExecute = result;
            e.Handled = true;
        }

        async void mailSendCmdBinding_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            string filename = this.oriExcelFile.Text;
            await _saveConfig();
            await SaveAsExcel(tNewFileName.Text);

            _busy.IsBusy = true;
            await Task.Run(() =>
            {
                this.Dispatcher.InvokeAsync(() =>
                {
                    _busy.BusyContent = "正在发送日志...";

                    Email email = new Email();
                    email.host = "smtp.gmail.com";
                    email.mailFrom = this.MailUser;
                    email.mailPwd = this.emailpwd.Password;
                    email.mailSubject = System.IO.Path.GetFileNameWithoutExtension(this.tNewFileName.Text) + " " + Person.PersonName;
                    email.mailToArray = this.MailTo.Split(';');

                    email.attachmentsPath = new string[] { this.tNewFileName.Text };

                    email.SendAsync(new System.Net.Mail.SendCompletedEventHandler((obj, ee) =>
                    {
                        string msg = ee.Error != null ? "发送失败:\r\n" + ee.Error.Message : "发送成功";
                        _busy.BusyContent = msg;
                        _busy.IsBusy = false;
                        MessageBox.Show(this, msg);
                    }));
                });
            });
        }

        async Task _saveConfig()
        {
            LastFileName = this.tNewFileName.Text;

            _busy.IsBusy = true;
            _busy.BusyContent = "正在保存配置文件...";

            await Task.Run(() =>
            {
                var conf = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                conf.AppSettings.Settings["lastfilename"].Value = LastFileName;
                conf.AppSettings.Settings["mailuser"].Value = MailUser;
                conf.AppSettings.Settings["mailto"].Value = MailTo;
                conf.AppSettings.Settings["contentcell"].Value = ContentCell;
                conf.Save();

            }).ContinueWith((x) =>
            {
                this.Dispatcher.Invoke(() => { _busy.IsBusy = false; });
            });

        }

        void mailSendCmdBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = (string.IsNullOrEmpty(this.tReceiver.Text) ||
                string.IsNullOrEmpty(this.emailUser.Text) ||
                string.IsNullOrEmpty(this.emailpwd.Password))
            ? false : true;
            e.Handled = true;
        }

        void readCmdBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = string.IsNullOrEmpty(this.g1.Text) || string.IsNullOrEmpty(this.oriExcelFile.Text) ? false : true;
            e.Handled = true;
        }

        void readCmdBinding_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            _busy.IsBusy = true;
            _busy.BusyContent = "正在读取文档...";

            string fname = this.oriExcelFile.Text;
            string cell = this.g1.Text;

            var t = Task.Run(() =>
            {
                using (ExcelUnit excel = new ExcelUnit(fname))
                {
                    Person.DiaryContent = excel.ReadCell(cell);
                }
            });

            t.GetAwaiter().OnCompleted(() =>
            {
                this.Dispatcher.Invoke(() =>
                {
                    _busy.IsBusy = false;
                });
            });

            e.Handled = true;
        }

        async Task ReadExcel(string excelFilename)
        {
            if (!System.IO.File.Exists(excelFilename)) return;
            _busy.BusyContent = "正在加载文档，请稍候...";
            _busy.IsBusy = true;

            await Task.Run(() =>
            {
                using (ExcelUnit excel = new ExcelUnit(excelFilename))
                {
                    this.Dispatcher.InvokeAsync(() =>
                    {
                        Person = excel.Read();

                    }).Completed += (o, oe) =>
                    {
                        _busy.IsBusy = false;

                        //自动设置为当前日期
                        Person.Date = System.DateTime.Now.ToShortDateString();


                    };
                }
            });

        }

        async Task SaveAsExcel(string filename)
        {
            _busy.BusyContent = "正在保存文档，请稍候...";
            _busy.IsBusy = true;

            string oriFilename = this.oriExcelFile.Text;

            await Task.Factory.StartNew(() =>
            {
                using (ExcelUnit excel = new ExcelUnit(oriFilename))
                {
                    this.Dispatcher.InvokeAsync(() =>
                    {
                        excel.SaveAs(Person, filename);

                    }).Completed += (o, oe) =>
                    {
                        _busy.IsBusy = false;
                    };
                }
            });
        }

        async void cmdBinding_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择日志模板文件";
            openFileDialog.Filter = "文件（.xls）|*.xls";//文件扩展名
            if ((bool)openFileDialog.ShowDialog().GetValueOrDefault())
            {
                this.oriExcelFile.Text = openFileDialog.FileName;
                await ReadExcel(this.oriExcelFile.Text);
            }
            e.Handled = true;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        #region 毛玻璃
        //[StructLayout(LayoutKind.Sequential)]
        //public struct MARGINS
        //{
        //    public int cxLeftWidth;
        //    public int cxRightWidth;
        //    public int cyTopHeight;
        //    public int cyBottomHeight;
        //};

        //[DllImport("DwmApi.dll")]
        //public static extern int DwmExtendFrameIntoClientArea(
        //    IntPtr hwnd,
        //    ref MARGINS pMarInset);

        //private void ExtendAeroGlass(Window window)
        //{
        //    try
        //    {
        //        // 为WPF程序获取窗口句柄
        //        IntPtr mainWindowPtr = new WindowInteropHelper(window).Handle;
        //        HwndSource mainWindowSrc = HwndSource.FromHwnd(mainWindowPtr);
        //        mainWindowSrc.CompositionTarget.BackgroundColor = Colors.Transparent;

        //        // 设置Margins
        //        MARGINS margins = new MARGINS();

        //        // 扩展Aero Glass
        //        margins.cxLeftWidth = -1;
        //        margins.cxRightWidth = -1;
        //        margins.cyTopHeight = -1;
        //        margins.cyBottomHeight = -1;

        //        int hr = DwmExtendFrameIntoClientArea(mainWindowSrc.Handle, ref margins);
        //        if (hr < 0)
        //        {
        //            MessageBox.Show("DwmExtendFrameIntoClientArea Failed");
        //        }
        //    }
        //    catch (DllNotFoundException)
        //    {
        //        Application.Current.MainWindow.Background = Brushes.White;
        //    }
        //}

        #endregion
        private void mainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            //this.Background = Brushes.Transparent;
            //ExtendAeroGlass(this);
        }
    }
}
