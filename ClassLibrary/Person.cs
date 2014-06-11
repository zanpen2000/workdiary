using System.ComponentModel;

namespace ClassLibrary
{
    public class Person : INotifyPropertyChanged
    {
        private string name;
        public string PersonName
        {
            get { return name; }
            set
            {
                name = value;
                if (PropertyChanged != null)
                {
                    PropertyChanged(this, new PropertyChangedEventArgs("PersonName"));
                }
            }
        }

        private string id;
        public string Id
        {
            get { return id; }
            set
            {
                id = value;
                if (PropertyChanged != null)
                {
                    PropertyChanged(this, new PropertyChangedEventArgs("Id"));
                }

            }
        }

        private string date;
        public string Date
        {
            get { return date; }
            set
            {
                date = value;
                if (PropertyChanged != null)
                {
                    PropertyChanged(this, new PropertyChangedEventArgs("Date"));
                }

            }
        }

        private string depart;
        public string Department
        {
            get { return depart; }
            set
            {
                depart = value;
                if (PropertyChanged != null)
                {
                    PropertyChanged(this, new PropertyChangedEventArgs("Department"));
                }
            }
        }

        private string company;
        public string Company
        {
            get { return company; }
            set
            {
                company = value;
                if (PropertyChanged != null)
                {
                    PropertyChanged(this, new PropertyChangedEventArgs("Company"));
                }
            }
        }

        private string diaryContent;
        public string DiaryContent
        {
            get { return diaryContent; }
            set
            {
                diaryContent = value;
                if (PropertyChanged != null)
                {
                    PropertyChanged(this, new PropertyChangedEventArgs("DiaryContent"));
                }
            }
        }


        public event PropertyChangedEventHandler PropertyChanged;
    }
}
