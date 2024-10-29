using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Configuration;
using AutoUploadBosa.ViewModels;
using System.Threading.Tasks;
using System.Net.Mail;
using Renci.SshNet;

namespace AutoUploadBosa
{
    public partial class Main : Form
    {
        private readonly string fromAddress = "SFCS_APServer@it.gemteks.com";
        public readonly string emailServer = "10.5.1.30";
        public string appPath = AppDomain.CurrentDomain.BaseDirectory;
        public string outFileTime = DateTime.Now.ToString("yyyyMMdd");
        DateTime theDate = DateTime.Now;
        public Main()
        {
            InitializeComponent();
        }
        private async void Main_Load(object sender, EventArgs e)
        {
            string sftpHost = ConfigurationManager.AppSettings.Get("sftpHost");
            int sftpPort = int.Parse(ConfigurationManager.AppSettings.Get("sftpPort"));
            string sftpUsername = ConfigurationManager.AppSettings.Get("sftpUsername");
            string sftpPassword = ConfigurationManager.AppSettings.Get("sftpPassword");
            string sftpNokia = ConfigurationManager.AppSettings.Get("sftpNokia");
            string localFilePathNokia = appPath.TrimEnd('\\') + ConfigurationManager.AppSettings.Get("localFilePath_Nokia");
            string sftpScc = ConfigurationManager.AppSettings.Get("sftpScc");
            string localFilePathScc = appPath.TrimEnd('\\') + ConfigurationManager.AppSettings.Get("localFilePath_Scc");
            using (var sftp = new SftpClient(sftpHost, sftpPort, sftpUsername, sftpPassword))
            {
                try
                {
                    if (!sftp.IsConnected)
                    {
                        sftp.Connect();
                    }

                    if (sftpNokia.Contains("NOKIA"))
                    {
                        try
                        {
                            await DownloadExcelFromNokia(sftp, sftpNokia, localFilePathNokia);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"The error downloading from Nokia: {ex.Message}");
                            if (!sftp.IsConnected)
                            {
                                sftp.Connect();
                            }
                        }
                    }

                    if (sftpScc.Contains("SCC"))
                    {
                        try
                        {
                            await DownloadExcelFromScc(sftp, sftpScc, localFilePathScc);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"The error downloading from SCC: {ex.Message}");
                            if (!sftp.IsConnected)
                            {
                                sftp.Connect();
                            }
                        }
                    }
                    try
                    {
                        await CopyFolder(
                            appPath.TrimEnd('\\') + ConfigurationManager.AppSettings.Get("localFilePath_Bosa"),
                            appPath.TrimEnd('\\') + ConfigurationManager.AppSettings.Get("localFilePath_Backup")
                        );
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"The error copying folder: {ex.Message}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"App can't connect to SFTP server: {ex.Message}");
                }
                finally
                {
                    if (sftp.IsConnected)
                    {
                        sftp.Disconnect();
                    }
                }
            }
            this.Close();
        }
        private static async Task CopyFolder(string sourceDir, string destDir)
        {
            Directory.CreateDirectory(destDir);
            var files = Directory.GetFiles(sourceDir);
            var directories = Directory.GetDirectories(sourceDir);
            var fileTasks = new List<Task>();
            foreach (var file in files)
            {
                string destFile = Path.Combine(destDir, Path.GetFileName(file));
                fileTasks.Add(Task.Run(() => File.Copy(file, destFile, true)));
            }
            await Task.WhenAll(fileTasks);
            var directoryTasks = new List<Task>();
            foreach (var subDir in directories)
            {
                string destSubDir = Path.Combine(destDir, Path.GetFileName(subDir));
                directoryTasks.Add(CopyFolder(subDir, destSubDir));
            }
            await Task.WhenAll(directoryTasks);
        }

        private async Task DownloadExcelFromNokia(SftpClient sftp, string remoteDirectory, string localDirectory)
        {
            List<string> listFile = new List<string>();
            try
            {
                if (!Directory.Exists(localDirectory))
                {
                    Directory.CreateDirectory(localDirectory);
                }
                if (!sftp.IsConnected)
                {
                    sftp.Connect();
                }
                var files = sftp.ListDirectory(remoteDirectory);
                foreach (var file in files)
                {
                    if (!file.IsRegularFile)
                    {
                        continue;
                    }
                    string remoteFilePath = file.FullName;
                    string localFilePath = Path.Combine(localDirectory, file.Name);
                    try
                    {
                        if (!sftp.IsConnected)
                        {
                            sftp.Connect();
                        }
                        using (var fileStream = new FileStream(localFilePath, FileMode.Create))
                        {
                            await Task.Run(() => sftp.DownloadFile(remoteFilePath, fileStream));
                        }
                        listFile.Add(localFilePath);
                        Bosa bosa = new Bosa();
                        await bosa.WriteDataBosaNokia(localFilePath);
                        if (!sftp.IsConnected)
                        {
                            sftp.Connect();
                        }
                        sftp.DeleteFile(remoteFilePath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"The error processing file '{file.Name}': {ex.Message}");
                        sftp.DeleteFile(remoteFilePath);
                    }
                }
                await AutoSendMail(string.Join(", ", listFile), "NOKIA");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"The error during SFTP operation: {ex.Message}");
            }
        }
        private async Task DownloadExcelFromScc(SftpClient sftp, string remoteDirectory, string localDirectory)
        {
            List<string> listFile = new List<string>();
            try
            {
                if (!Directory.Exists(localDirectory))
                {
                    Directory.CreateDirectory(localDirectory);
                }
                if (!sftp.IsConnected)
                {
                    sftp.Connect();
                }
                var files = sftp.ListDirectory(remoteDirectory);
                foreach (var file in files)
                {
                    if (!file.IsRegularFile)
                    {
                        continue;
                    }
                    string remoteFilePath = file.FullName;
                    string localFilePath = Path.Combine(localDirectory, file.Name);
                    try
                    {
                        if (!sftp.IsConnected)
                        {
                            sftp.Connect();
                        }
                        using (var fileStream = new FileStream(localFilePath, FileMode.Create))
                        {
                            await Task.Run(() => sftp.DownloadFile(remoteFilePath, fileStream));
                        }
                        listFile.Add(localFilePath);
                        Bosa bosa = new Bosa();
                        await bosa.WriteDataBosaScc(localFilePath);
                        if (!sftp.IsConnected)
                        {
                            sftp.Connect();
                        }
                        sftp.DeleteFile(remoteFilePath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"The error processing file '{file.Name}': {ex.Message}");
                        sftp.DeleteFile(remoteFilePath);
                    }
                }
                await AutoSendMail(string.Join(", ", listFile), "SCC");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"The error during SFTP operation: {ex.Message}");
            }
        }

        public async Task AutoSendMail(string _listFile, string _dataBaseName)
        {
            string FilePath = appPath + "File\\MailSendTo.txt";
            string emailSubject = "", emailBody = "";
            if (_dataBaseName.Contains("NOKIA"))
            {
                emailSubject = "The uploading bosa data for NOKIA database " + theDate.ToString("yyyy/MM/dd");
                emailBody = "Dear all, \n";
                if (!string.IsNullOrEmpty(_listFile))
                {
                    emailBody += "The BOSA in the NOKIA database has been uploaded successfully.\n" + ($"{_listFile} file") + "\n Please check the in below. \n";
                }
                else
                {
                    emailBody += "The BOSA in the NOKIA database doesn't have any uploaded data" + "\n Please check the information in below. \n";
                }
            }
            if (_dataBaseName.Contains("SCC"))
            {
                emailSubject = "The uploading bosa data for SCC database " + theDate.ToString("yyyy/MM/dd");
                emailBody = "Dear all, \n";
                if (!string.IsNullOrEmpty(_listFile))
                {
                    emailBody += "The BOSA in the SCC database has been uploaded successfully.\n" + ($"{_listFile} file") + "\n Please check the in below. \n";
                }
                else
                {
                    emailBody += "The BOSA in the SCC database doesn't have any uploaded data" + "\n Please check the information in below. \n";
                }
            }
            emailBody += "Thanks & Best Regard! \n";
            emailBody += "----- \n";
            emailBody += "Tel: \n";
            emailBody += "Gemtek Vietnam Co.,Ltd. \n";
            try
            {
                string fileEmails = appPath + "File\\MailSendTo.txt";
                string reportToAddress = await GetListEmail(fileEmails);
                MailMessage theMessage = new MailMessage(fromAddress, reportToAddress);
                theMessage.Subject = emailSubject;
                theMessage.Body = emailBody;
                theMessage.IsBodyHtml = false;
                theMessage.SubjectEncoding = System.Text.Encoding.GetEncoding("BIG5");
                theMessage.BodyEncoding = System.Text.Encoding.GetEncoding("BIG5");
                SmtpClient theSmtpServer = new SmtpClient();
                theSmtpServer.Credentials = new System.Net.NetworkCredential("cs_sfcs", "gemtek12345");
                theSmtpServer.Host = emailServer;
                await theSmtpServer.SendMailAsync(theMessage);
            }
            catch (Exception ex)
            {
                MessageBox.Show("error: " + ex.Message);
            }
        }
        private async Task<string> GetListEmail(string strFile)
        {
            string emailList = "";
            using (FileStream fs = new FileStream(strFile, FileMode.Open, FileAccess.Read))
            using (StreamReader myStreamReader = new StreamReader(fs))
            {
                try
                {
                    while (myStreamReader.Peek() != -1)
                    {
                        emailList += await myStreamReader.ReadLineAsync() + ",";
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"The file error occurred: {ex.Message}");
                }
            }
            if (emailList.Length > 0)
            {
                emailList = emailList.Substring(0, emailList.Length - 1);
            }
            return emailList;
        }
    }
}
