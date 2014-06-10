using ClassLibrary;
using Microsoft.Win32;
using NetOffice.ExcelApi;
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
using System.Configuration;
using System.ComponentModel;
using System.Windows.Threading;

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
            CommandBinding readExcelCmdBinding = new CommandBinding();
            readExcelCmdBinding.Command = ReadExcelCommand;
            readExcelCmdBinding.Executed += readExcelCmdBinding_Executed;
            readExcelCmdBinding.CanExecute += readExcelCmdBinding_CanExecute;

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

            this.mainWindow.CommandBindings.Add(browseCmdBinding);
            this.mainWindow.CommandBindings.Add(readCmdBinding);
            this.mainWindow.CommandBindings.Add(mailSendCmdBinding);

            this.oriExcelFile.TextChanged += (o, e) =>
            {
                if (System.IO.File.Exists(this.oriExcelFile.Text))
                {
                    ReadExcel(this.oriExcelFile.Text);
                }
            };
        }



        void readExcelCmdBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = System.IO.File.Exists(this.oriExcelFile.Text) ? true : false;
            e.Handled = true;
        }

        async void readExcelCmdBinding_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            await ReadExcel(this.oriExcelFile.Text);
            e.Handled = true;
        }

        async void mailSendCmdBinding_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            string filename = this.oriExcelFile.Text;
            await _saveConfig();
            await SaveAsExcel(tNewFileName.Text);

            _busy.IsBusy = true;
            await Task.Factory.StartNew(() =>
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

            await Task.Factory.StartNew(() =>
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
                string.IsNullOrEmpty(this.emailpwd.Password)) ||
                oriExcelFile.Text.Trim().Equals(tNewFileName.Text.Trim())
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

            await Task.Factory.StartNew(() =>
            {
                using (ExcelUnit excel = new ExcelUnit(excelFilename))
                {
                    this.Dispatcher.InvokeAsync(() =>
                    {
                        Person = excel.Read();

                    }).Completed += (o, oe) =>
                    {
                        _busy.IsBusy = false;
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
    }
}
