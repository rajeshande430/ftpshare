using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
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
using System.Windows.Shapes;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace FTP_Share
{
    /// <summary>
    /// Interaction logic for ShareFTPForm.xaml
    /// </summary>
    public partial class ShareFTPForm : Window
    {
        public static ShareFTPForm Object;
        public static Item Item = null;

        public ShareFTPForm()
        {
            InitializeComponent();

            cmbx_projectnames.ItemsSource = SharepointHelper.GetSubFolderNames();
            Object = this;
        }

        private async void OnShareToFTP(object sender, RoutedEventArgs e)
        {
            if ((cmbx_projectnames.SelectedItem == null))
            {
                txt_shareError.Foreground = new SolidColorBrush(Colors.Red);
                txt_shareError.Text = "Please select the project folder.";
                return;
            }

            button_shareFtp.IsEnabled = false;
            button_selectFolder.IsEnabled = false;
            img_loading.Visibility = Visibility.Visible;
            txt_shareError.Text = string.Empty;
            txt_shareEmail.Visibility = Visibility.Hidden;

            if ((string.IsNullOrEmpty(txt_folderPath.Text) || string.IsNullOrEmpty(cmbx_projectnames.SelectedItem.ToString())))
            {
                txt_shareError.Foreground = new SolidColorBrush(Colors.Red);
                txt_shareError.Text = "Please select file and project where you wish to share it.";
                return;
            }

            var selectProject = cmbx_projectnames.SelectedItem.ToString();

            if (SharepointHelper.IsFolderExistSP(cmbx_projectnames.SelectedItem.ToString(), Item.Name))
            {
                var result = System.Windows.MessageBox.Show($"'{Item.Name}' folder already exist. Do you want to override it?", "Important", MessageBoxButton.YesNo, MessageBoxImage.Information);

                if (result == MessageBoxResult.No) return;
            }

            await SharepointHelper.UploadFolderToSharePoint(Item, selectProject);
            var hyperlink = SharepointHelper.GetFolderHyperLink(Item.Folder); ;
            SharepointHelper.InsertEnquiryToSharepoint(txt_folderPath.Text, Item.Folder, hyperlink);
            txt_shareHyperlink.Text = hyperlink;

            txt_uploadingsize.Visibility = Visibility.Hidden;
            img_loading.Visibility = Visibility.Hidden;

            txt_shareEmail.Visibility = Visibility.Visible;
            //txt_folderPath.Text = string.Empty;

            if (!string.IsNullOrEmpty(txt_shareHyperlink.Text))
            {
                txt_shareError.Foreground = new SolidColorBrush(Colors.Green);
                txt_shareError.Text = "The shared link will get expired after 14 Days.\n\t   Shared Successfully!";
            }

            button_selectFolder.IsEnabled = true;
            btn_copyclipboard.IsEnabled = true;

        }

        private void OnSelectFolder(object sender, RoutedEventArgs e)
        {
            txt_shareEmail.Visibility = Visibility.Hidden;
            txt_shareHyperlink.Text = string.Empty;
            txt_shareError.Text = string.Empty;


            var fbd = new Util.FolderSelectDialog();
            fbd.Title = "Select folder to upload";
            if (fbd.ShowDialog(IntPtr.Zero))
            {
                var directoryPath = fbd.FileName;
    
                Item = new Item();
                Item.FullPath = directoryPath;
                Item.Type = ItemType.Folder;
                Item.RelativePath = "";
                Item.Name = new System.IO.DirectoryInfo(directoryPath).Name;

                TraverseInDirectory(Item);
                txt_folderPath.Text = Item.FullPath;

                button_shareFtp.IsEnabled = true;
                btn_copyclipboard.IsEnabled = false;
            }

        }

        private void OnOpenOutlook(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Outlook.Application oApp = new Outlook.Application();
                Outlook._MailItem oMailItem = (Outlook._MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                oMailItem.Body = txt_shareHyperlink.Text;
                oMailItem.Display(true);
            }
            catch (Exception)
            {

            }
        }

        private void CopytoClipBoard(object sender, RoutedEventArgs e)
        {
            try
            {
                Clipboard.SetText(txt_shareHyperlink.Text);
                txt_shareError.Text = "Copied Successfully.";
            }
            catch (Exception)
            {
            }
        }

        public static void TraverseInDirectory(Item item)
        {

            foreach (string subdirpath in Directory.GetDirectories(item.FullPath))
            {
                var subitem = new Item();
                subitem.FullPath = subdirpath;
                subitem.Type = ItemType.Folder;
                subitem.Name = new DirectoryInfo(subdirpath).Name;
                subitem.RelativePath = item.RelativePath + item.Name + "/";
                item.Items.Add(subitem);

                TraverseInDirectory(subitem);
            }

            foreach (string filepath in Directory.GetFiles(item.FullPath))
            {
                var file = new Item();
                file.Type = ItemType.File;
                file.FullPath = filepath;
                file.RelativePath = item.RelativePath + item.Name + "/";
                file.Name = new FileInfo(filepath).Name;
                item.Items.Add(file);
            }


        }



        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            //do my stuff before closing
            var result = System.Windows.MessageBox.Show("Are you sure you want to close the application?", "Warning", MessageBoxButton.YesNo, MessageBoxImage.Information);

            if (result == MessageBoxResult.Yes)
            {
                base.OnClosing(e);
            }
            else
            {
                e.Cancel = true;
            }
        }
    }
}
