using Microsoft.Win32;
using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Windows;
using Spire.Doc;
using ImapX;


namespace WordLekcia
{
    public partial class Window3 : Window
    {
        private string selectedFilePath;

        public Window3()
        {
            InitializeComponent();
        }

        private void SelectFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Word files (*.docx)|*.docx|All files (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                selectedFilePath = openFileDialog.FileName;
                MessageBox.Show("Выбран файл: " + selectedFilePath);
            }
        }

        private void SendButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFilePath))
            {
                MessageBox.Show("Пожалуйста, выберите файл для отправки.");
                return;
            }

            try
            {
                string tempFilePath = Path.GetTempFileName();

                Document doc = new Document();
                doc.LoadFromFile(selectedFilePath);

                doc.SaveToFile(tempFilePath, FileFormat.Docx);

                MailMessage messagev = new MailMessage(From.Text, To.Text, Subject.Text, "Пожалуйста, найдите прикрепленный документ.");

                Attachment attachment = new Attachment(tempFilePath);
                messagev.Attachments.Add(attachment);

                SmtpClient smtpClient = new SmtpClient("smtp.mail.ru", 587);
                smtpClient.Credentials = new NetworkCredential(From.Text, Pass.Password);
                smtpClient.EnableSsl = true;
                smtpClient.Send(messagev);

                File.Delete(tempFilePath);

                MessageBox.Show("Сообщение отправлено успешно.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка отправки сообщения: " + ex.Message);
            }
        }
    }
}
