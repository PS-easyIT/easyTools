using System;
using System.ComponentModel;
using System.Windows;

namespace EasyWSUS
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = new ViewModelMain();
        }
    }

    public class ViewModelMain : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _appName = "WSUS Manager";
        public string AppName
        {
            get { return _appName; }
            set
            {
                _appName = value;
                OnPropertyChanged(nameof(AppName));
            }
        }

        private string _footerText = "© 2023 EASY IT - Alle Rechte vorbehalten";
        public string FooterText
        {
            get { return _footerText; }
            set
            {
                _footerText = value;
                OnPropertyChanged(nameof(FooterText));
            }
        }

        private string _footerWebsite = "www.easy-it.de";
        public string FooterWebsite
        {
            get { return _footerWebsite; }
            set
            {
                _footerWebsite = value;
                OnPropertyChanged(nameof(FooterWebsite));
            }
        }

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
