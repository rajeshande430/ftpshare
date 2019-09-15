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

namespace FTP_Share
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            GetLastUserCredentials();
        }

        private void GetLastUserCredentials()
        {
            txtbx_email.Text = Properties.Settings.Default.username;
            txtbx_password.Password = Properties.Settings.Default.password;
        }

        private Task<string> LoginTo365Async(string email, string password, string url = "https://archcorp365.sharepoint.com/sites/archcorpftp/")
        {
            return Task.Run(() => 
            {
                if (string.IsNullOrEmpty(email) || string.IsNullOrEmpty(password))
                {
                    return "Please type correct email and password";
                }

                try
                {
                    SharepointHelper.Login(email, password, url);
                }
                catch (Exception ex)
                {
                    return ex.Message;
                    
                }


                return "";
            });
        }

        private async void OnLoginO365(object sender, RoutedEventArgs e)
        {
            if (!SharepointHelper.CheckIfTheUserIsInTheNetwork())
            {
                System.Windows.MessageBox.Show( "Sorry you can't use any network outside Archcorp's domain. To use the application, please connect to archcorp network or get in touch with IT admins. Thank you",
                    "Access Denied", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                button_login.IsEnabled = true;
                return;
            }


            button_login.IsEnabled = false;

            txt_loginerror.Text = string.Empty;
            string email = txtbx_email.Text;
            string password = txtbx_password.Password;
            string url = "https://archcorp365.sharepoint.com/sites/archcorpftp/";

            txt_loginerror.Text = await LoginTo365Async(email, password, url);
            button_login.IsEnabled = true;

            // If there is some error then don't return
            if (!String.IsNullOrEmpty(txt_loginerror.Text)) return;

            Properties.Settings.Default.username = txtbx_email.Text;
            Properties.Settings.Default.password = txtbx_password.Password;
            Properties.Settings.Default.Save();

            new ShareFTPForm().Show();
            this.Close();

        }

        private void OnKeyDownHandler(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Return)
            {
                button_login.IsEnabled = false;
                //textBlock1.Text = "You Entered: " + textBox1.Text;
                OnLoginO365(sender, e);
            }
        }

        private void OnForgetPassword(object sender, RoutedEventArgs e)
        {
            System.Windows.MessageBox.Show("Please contact the IT Administrator for Login in case you've forgot the password.", "Forgot Password", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
