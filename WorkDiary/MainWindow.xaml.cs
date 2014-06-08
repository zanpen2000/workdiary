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
        RoutedCommand SaveAsCommand = new RoutedCommand("SaveAs", typeof(MainWindow));
        RoutedCommand MailSendCommand = new RoutedCommand("SendMail", typeof(MainWindow));
        RoutedCommand ReadExcelCommand = new RoutedCommand("ReadExcel", typeof(MainWindow));
        RoutedCommand SaveConfigCommand = new RoutedCommand("SaveConfig", typeof(MainWindow));

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

            DataContext = this;

            this.Dispatcher.BeginInvoke((System.Action)delegate
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
            }, null);


        }

        private void SetCommandBinding()
        {
            CommandBinding saveCfgCmdBinding = new CommandBinding();
            saveCfgCmdBinding.Command = SaveConfigCommand;
            saveCfgCmdBinding.Executed += saveCfgCmdBinding_Executed;
            saveCfgCmdBinding.CanExecute += saveCfgCmdBinding_CanExecute;

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


            this.btnSaveAs.Command = SaveAsCommand;
            this.SaveAsCommand.InputGestures.Add(new KeyGesture(Key.S, ModifierKeys.Alt));
            this.btnSaveAs.CommandTarget = this.tNewFileName;

            CommandBinding saveAsBinding = new CommandBinding();
            saveAsBinding.Command = SaveAsCommand;
            saveAsBinding.Executed += saveAsBinding_Executed;
            saveAsBinding.CanExecute += saveAsBinding_CanExecute;

            this.btnSend.Command = MailSendCommand;
            MailSendCommand.InputGestures.Add(new KeyGesture(Key.M, ModifierKeys.Alt));
            this.btnSend.CommandTarget = this.tReceiver;
            CommandBinding mailSendCmdBinding = new CommandBinding();
            mailSendCmdBinding.Command = MailSendCommand;
            mailSendCmdBinding.CanExecute += mailSendCmdBinding_CanExecute;
            mailSendCmdBinding.Executed += mailSendCmdBinding_Executed;

            this.mainWindow.CommandBindings.Add(browseCmdBinding);
            this.mainWindow.CommandBindings.Add(readCmdBinding);
            this.mainWindow.CommandBindings.Add(saveAsBinding);
            this.mainWindow.CommandBindings.Add(mailSendCmdBinding);

            this.oriExcelFile.TextChanged += (o, e) =>
            {
                if (System.IO.File.Exists(this.oriExcelFile.Text))
                {
                    ReadExcel(this.oriExcelFile.Text);
                }
            };
        }

        void saveCfgCmdBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            throw new NotImplementedException();
        }

        void saveCfgCmdBinding_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            throw new NotImplementedException();
        }

        void readExcelCmdBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = System.IO.File.Exists(this.oriExcelFile.Text) ? true : false;
            e.Handled = true;
        }

        void readExcelCmdBinding_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            ReadExcel(this.oriExcelFile.Text);
            e.Handled = true;
        }

        void mailSendCmdBinding_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            BackgroundWorker bw = new BackgroundWorker();

            bw.DoWork += (x, y) =>
            {
                this.Dispatcher.Invoke(() =>
                {
                    _busy.BusyContent = "正在保存配置文件...";
                    SaveConfig();

                    _busy.BusyContent = "正在保存日志...";

                    using (ExcelUnit excel = new ExcelUnit(this.oriExcelFile.Text))
                    {
                        excel.SaveAs(Person, tNewFileName.Text);
                    }

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
                        this.Dispatcher.InvokeAsync(() =>
                        {
                            _busy.IsBusy = false;
                            string msg = ee.Error != null ? "发送失败:\r\n" + ee.Error.Message : "发送成功";
                            MessageBox.Show(this, msg);
                        });
                    }));
                });
            };

            bw.RunWorkerAsync();
            _busy.IsBusy = true;
        }

        private void SaveConfig()
        {
            LastFileName = this.tNewFileName.Text;

            var conf = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            conf.AppSettings.Settings["lastfilename"].Value = LastFileName;
            conf.AppSettings.Settings["mailuser"].Value = MailUser;
            conf.AppSettings.Settings["mailto"].Value = MailTo;
            conf.AppSettings.Settings["contentcell"].Value = ContentCell;
            conf.Save();
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

        void saveAsBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = !string.IsNullOrEmpty(this.tNewFileName.Text) &&
                !this.tNewFileName.Text.Equals(this.oriExcelFile.Text);
            e.Handled = true;
        }

        void saveAsBinding_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            //保存信息，个人信息，日志内容，时间
            SaveFileDialog sDialog = new SaveFileDialog();
            sDialog.Title = "选择保存路径";
            sDialog.Filter = "文件（.xls）|*.xls";//文件扩展名
            sDialog.FileName = this.tNewFileName.Text;
            if ((bool)sDialog.ShowDialog().GetValueOrDefault())
            {
                using (ExcelUnit excel = new ExcelUnit(this.oriExcelFile.Text))
                {
                    string msg = excel.SaveAs(Person, sDialog.FileName) ? "保存成功" : "保存失败";
                    MessageBox.Show(this, msg);
                }
                SaveConfig();
            }
            e.Handled = true;
        }

        void readCmdBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = string.IsNullOrEmpty(this.g1.Text) ? false : true;
            e.Handled = true;
        }

        void readCmdBinding_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            using (ExcelUnit excel = new ExcelUnit(this.oriExcelFile.Text))
            {
                Person.DiaryContent = excel.ReadCell(this.g1.Text);
            }

            e.Handled = true;
        }

        void ReadExcel(string excelFilename)
        {
            if (!System.IO.File.Exists(excelFilename)) return;

            this.Dispatcher.InvokeAsync(() =>
            {
                using (ExcelUnit excel = new ExcelUnit(excelFilename))
                {
                    Person = excel.Read();
                }

            });


        }

        void cmdBinding_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "选择日志模板文件";
            openFileDialog.Filter = "文件（.xls）|*.xls";//文件扩展名
            if ((bool)openFileDialog.ShowDialog().GetValueOrDefault())
            {
                //这里为什么不显示

                _busy.BusyContent = "正在加载文档，请稍候...";
                _busy.IsBusy = true;

                this.Dispatcher.InvokeAsync(() =>
                {
                    this.oriExcelFile.Text = openFileDialog.FileName;
                    ReadExcel(this.oriExcelFile.Text);

                }, DispatcherPriority.Background).Completed += (o, oe) =>
                {
                    _busy.IsBusy = false;
                };
            }
            e.Handled = true;
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
