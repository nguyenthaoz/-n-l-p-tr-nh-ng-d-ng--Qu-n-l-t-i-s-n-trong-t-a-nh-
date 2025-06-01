using PhanMemQuanLyTaiSan.view;
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

namespace PhanMemQuanLyTaiSan
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private void LoginButton_Click(object sender, RoutedEventArgs e)
        {
            string username = taikhoan.Text;
            string password = matkhau.Password;

            if (IsValidUser(username, password))
            {
                HomePage homepage = new HomePage();
                homepage.Show();
                this.Close();
               
            }
            else
            {
                MessageBox.Show("Tên đăng nhập hoặc mật khẩu không đúng.");
            }
        }

        private bool IsValidUser(string username, string password)
        {
            return (username == "admin" && password == "123") ||
                   (username == "nvkythuat1" && password == "123") ||
                   (username == "nvkythuat2" && password == "123");
        }
    }
}
