# PolicyPrivacy12
ZigZag Rainbow 3D

**Terms & Conditions**

By downloading or using the app, these terms will automatically apply to you – you should make sure therefore that you read them carefully before using the app. You’re not allowed to copy or modify the app, any part of the app, or our trademarks in any way. You’re not allowed to attempt to extract the source code of the app, and you also shouldn’t try to translate the app into other languages or make derivative versions. The app itself, and all the trademarks, copyright, database rights, and other intellectual property rights related to it, still belong to Justtouch Games.

Justtouch Games is committed to ensuring that the app is as useful and efficient as possible. For that reason, we reserve the right to make changes to the app or to charge for its services, at any time and for any reason. We will never charge you for the app or its services without making it very clear to you exactly what you’re paying for.

The ZigZag Rainbow 3D app stores and processes personal data that you have provided to us, to provide my Service. It’s your responsibility to keep your phone and access to the app secure. We therefore recommend that you do not jailbreak or root your phone, which is the process of removing software restrictions and limitations imposed by the official operating system of your device. It could make your phone vulnerable to malware/viruses/malicious programs, compromise your phone’s security features and it could mean that the ZigZag Rainbow 3D app won’t work properly or at all.

The app does use third-party services that declare their Terms and Conditions.

Link to Terms and Conditions of third-party service providers used by the app

*   [Google Play Services](https://policies.google.com/terms)
*   [AdMob](https://developers.google.com/admob/terms)
*   [Google Analytics for Firebase](https://firebase.google.com/terms/analytics)
*   [Firebase Crashlytics](https://firebase.google.com/terms/crashlytics)
*   [Unity](https://unity3d.com/legal/terms-of-service)

You should be aware that there are certain things that Justtouch Games will not take responsibility for. Certain functions of the app will require the app to have an active internet connection. The connection can be Wi-Fi or provided by your mobile network provider, but Justtouch Games cannot take responsibility for the app not working at full functionality if you don’t have access to Wi-Fi, and you don’t have any of your data allowance left.

If you’re using the app outside of an area with Wi-Fi, you should remember that the terms of the agreement with your mobile network provider will still apply. As a result, you may be charged by your mobile provider for the cost of data for the duration of the connection while accessing the app, or other third-party charges. In using the app, you’re accepting responsibility for any such charges, including roaming data charges if you use the app outside of your home territory (i.e. region or country) without turning off data roaming. If you are not the bill payer for the device on which you’re using the app, please be aware that we assume that you have received permission from the bill payer for using the app.

Along the same lines, Justtouch Games cannot always take responsibility for the way you use the app i.e. You need to make sure that your device stays charged – if it runs out of battery and you can’t turn it on to avail the Service, Justtouch Games cannot accept responsibility.

With respect to Justtouch Games’s responsibility for your use of the app, when you’re using the app, it’s important to bear in mind that although we endeavor to ensure that it is updated and correct at all times, we do rely on third parties to provide information to us so that we can make it available to you. Justtouch Games accepts no liability for any loss, direct or indirect, you experience as a result of relying wholly on this functionality of the app.

At some point, we may wish to update the app. The app is currently available on iOS – the requirements for the system(and for any additional systems we decide to extend the availability of the app to) may change, and you’ll need to download the updates if you want to keep using the app. Justtouch Games does not promise that it will always update the app so that it is relevant to you and/or works with the iOS version that you have installed on your device. However, you promise to always accept updates to the application when offered to you, We may also wish to stop providing the app, and may terminate use of it at any time without giving notice of termination to you. Unless we tell you otherwise, upon any termination, (a) the rights and licenses granted to you in these terms will end; (b) you must stop using the app, and (if needed) delete it from your device.

**Changes to This Terms and Conditions**

I may update our Terms and Conditions from time to time. Thus, you are advised to review this page periodically for any changes. I will notify you of any changes by posting the new Terms and Conditions on this page.

These terms and conditions are effective as of 2021-11-30

**Contact Us**

If you have any questions or suggestions about my Terms and Conditions, do not hesitate to contact me at justtouch@gmail.com.





// DataAccess/Country.cs
namespace ExcelFileProcessor.DataAccess
{
    public class Country
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Code { get; set; }
    }
}



// DataAccess/DatabaseManager.cs
using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace ExcelFileProcessor.DataAccess
{
    public class DatabaseManager
    {
        private string connectionString = "Data Source=DESKTOP-OPQQL1L\\SQLEXPRESS;Initial Catalog=DenemeExcelVT;Integrated Security=True";

        public void InsertData(string mesken, string ulke, string saat, string sehir, string hayvan, string oyun)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Önce ülkeyi kontrol et ve gerekirse ekleyerek ülke ID'sini al
                int ulkeId = GetOrInsertCountryId(ulke, connection);

                string insertQuery = "INSERT INTO DenemeExcelVT.dbo.KayitExcel (Mesken, UlkeId, Saat, Sehir, Hayvan, Oyun) VALUES (@Mesken, @UlkeId, @Saat, @Sehir, @Hayvan, @Oyun)";

                using (SqlCommand command = new SqlCommand(insertQuery, connection))
                {
                    command.Parameters.AddWithValue("@Mesken", mesken);
                    command.Parameters.AddWithValue("@UlkeId", ulkeId);
                    command.Parameters.AddWithValue("@Saat", string.IsNullOrEmpty(saat) ? (object)DBNull.Value : saat);
                    command.Parameters.AddWithValue("@Sehir", sehir);
                    command.Parameters.AddWithValue("@Hayvan", hayvan);
                    command.Parameters.AddWithValue("@Oyun", oyun);

                    command.ExecuteNonQuery();
                }
            }
        }

        private int GetOrInsertCountryId(string ulke, SqlConnection connection)
        {
            // Ülkeyi kontrol et
            int ulkeId = GetCountryId(ulke, connection);

            // Eğer ülke listede yoksa, ekleyerek ID'sini al
            if (ulkeId == 0)
            {
                ulkeId = InsertCountryAndGetId(ulke, connection);
            }

            return ulkeId;
        }

        private int GetCountryId(string ulke, SqlConnection connection)
        {
            string selectQuery = "SELECT Id FROM ulkeler WHERE Name = @Ulke";

            using (SqlCommand command = new SqlCommand(selectQuery, connection))
            {
                command.Parameters.AddWithValue("@Ulke", ulke);

                var result = command.ExecuteScalar();

                return (result != null && result != DBNull.Value) ? Convert.ToInt32(result) : 0;
            }
        }

        private int InsertCountryAndGetId(string ulke, SqlConnection connection)
        {
            string insertQuery = "INSERT INTO ulkeler (Name) OUTPUT INSERTED.ID VALUES (@Ulke)";

            using (SqlCommand command = new SqlCommand(insertQuery, connection))
            {
                command.Parameters.AddWithValue("@Ulke", ulke);

                return (int)command.ExecuteScalar();
            }
        }
    }
}



// DataAccess/ErrorLogManager.cs
using System;
using System.IO;

namespace ExcelFileProcessor.DataAccess
{
    public class ErrorLogManager
    {
        private readonly string errorLogPath;

        public ErrorLogManager()
        {
            // Hata dosyalarının ana dizini
            string baseDirectory = @"D:\HedefExcel\hatalilar\";

            // Yıl, ay ve gün bilgilerini al
            string year = DateTime.Now.Year.ToString();
            string month = DateTime.Now.Month.ToString("00");
            string day = DateTime.Now.Day.ToString("00");

            // Hata dosyasının tam dizini
            errorLogPath = Path.Combine(baseDirectory, year, month, day);

            // Eğer dosya dizini yoksa oluştur
            if (!Directory.Exists(errorLogPath))
            {
                Directory.CreateDirectory(errorLogPath);
            }
        }

        public void LogError(string errorMessage)
        {
            // Hata mesajını dosyaya yaz
            string errorFilePath = Path.Combine(errorLogPath, "error_log.txt");
            File.AppendAllText(errorFilePath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {errorMessage}\n");
        }
    }
}








// Forms/Form1.cs
using System;
using System.IO;
using System.Timers;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFileProcessor.Forms
{
    public partial class Form1 : Form
    {
        private string kaynakDizin = @"D:\KaynakExcel";
        private string hedefDizin = @"D:\HedefExcel";
        private string hatalilarDizin = @"D:\HedefExcel\hatalilar";

        private System.Timers.Timer timer;
        private Queue<string> dosyaKuyrugu = new Queue<string>();
        private object kuyrukLock = new object();
        private bool islemDevamEdiyor = false;

        private DataAccess.DatabaseManager databaseManager = new DataAccess.DatabaseManager();
        private DataAccess.ErrorLogManager errorLogManager = new DataAccess.ErrorLogManager();

        public Form1()
        {
            InitializeComponent();

            timer = new System.Timers.Timer
            {
                Interval = 5000,
                AutoReset = true,
                Enabled = false
            };

            timer.Elapsed += Timer_Elapsed;
        }

        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            KontrolEtVeTasi();
        }

        private void KontrolEtVeTasi()
        {
            lock (kuyrukLock)
            {
                if (islemDevamEdiyor)
                    return;

                string[] excelDosyalari = Directory.GetFiles(kaynakDizin, "*.xlsx");
                foreach (string excelDosyasi in excelDosyalari)
                {
                    dosyaKuyrugu.Enqueue(excelDosyasi);
                }

                islemDevamEdiyor = true;
            }

            while (dosyaKuyrugu.Count > 0)
            {
                string excelDosyasi;
                lock (kuyrukLock)
                {
                    excelDosyasi = dosyaKuyrugu.Dequeue();
                }

                string dosyaAdi = Path.GetFileName(excelDosyasi);
                string hedefDosyaYolu = Path.Combine(hedefDizin, dosyaAdi);

                try
                {
                    File.Move(excelDosyasi, hedefDosyaYolu);

                    LogMesaji($"Dosya taşındı: {dosyaAdi}");

                    // Veritabanına ekle
                    ImportAndInsertToDatabase(hedefDosyaYolu);
                }
                catch (Exception ex)
                {
                    LogMesaji($"Hata oluştu: {ex.Message}");
                    LogErrorToFile(ex, dosyaAdi);

                    // Hata durumunda Excel dosyasını hatalilar dizinine taşı
                    MoveToErrorFolder(excelDosyasi);
                }
            }

            lock (kuyrukLock)
            {
                dosyaKuyrugu.Clear();
            }

            islemDevamEdiyor = false;
        }

        private void ImportAndInsertToDatabase(string excelDosyaYolu)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(excelDosyaYolu);
            Excel.Worksheet excelWorksheet = excelWorkbook.Sheets[1];

            try
            {
                for (int row = 2; row <= excelWorksheet.UsedRange.Rows.Count; row++)
                {
                    string mesken = excelWorksheet.Cells[row, 1].Value2?.ToString();
                    string ulke = excelWorksheet.Cells[row, 2].Value2?.ToString();
                    string saat = excelWorksheet.Cells[row, 4].Value2?.ToString();
                    string sehir = excelWorksheet.Cells[row, 5].Value2?.ToString();
                    string hayvan = excelWorksheet.Cells[row, 6].Value2?.ToString();
                    string oyun = excelWorksheet.Cells[row, 8].Value2?.ToString();

                    databaseManager.InsertData(mesken, ulke, saat, sehir, hayvan, oyun);
                }
            }
            catch (Exception ex)
            {
                LogMesaji($"Veritabanına ekleme hatası: {ex.Message}");
                LogErrorToFile(ex, excelDosyaYolu);

                // Hata durumunda Excel dosyasını hatalilar dizinine taşı
                MoveToErrorFolder(excelDosyaYolu);
            }
            finally
            {
                excelWorkbook.Close(false);
                excelApp.Quit();
            }
        }

        private void MoveToErrorFolder(string excelDosyaYolu)
        {
            try
            {
                string dosyaAdi = Path.GetFileName(excelDosyaYolu);

                // Yıl, ay ve gün bilgilerini al
                string year = DateTime.Now.Year.ToString();
                string month = DateTime.Now.Month.ToString("00");
                string day = DateTime.Now.Day.ToString("00");

                // Hatalilar dizininin tam dizini
                string hatalilarDosyaYolu = Path.Combine(hatalilarDizin, year, month, day, dosyaAdi);

                // Eğer dosya dizini yoksa oluştur
                string hatalilarDizinPath = Path.Combine(hatalilarDizin, year, month, day);
                if (!Directory.Exists(hatalilarDizinPath))
                {
                    Directory.CreateDirectory(hatalilarDizinPath);
                }

                File.Move(excelDosyaYolu, hatalilarDosyaYolu);
                LogMesaji($"Dosya hatalilar dizinine taşındı: {dosyaAdi}");
            }
            catch (Exception ex)
            {
                LogMesaji($"Hata dosya taşıma hatası: {ex.Message}");
            }
        }

        private void LogMesaji(string mesaj)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => LogMesaji(mesaj)));
            }
            else
            {
                listBoxLog.Items.Add($"{DateTime.Now:HH:mm:ss} - {mesaj}");
                listBoxLog.SelectedIndex = listBoxLog.Items.Count - 1;
            }
        }

        private void LogErrorToFile(Exception ex, string dosyaAdi)
        {
            string errorMessage = $"Hata dosya adı: {dosyaAdi}, Hata: {ex.Message}";
            errorLogManager.LogError(errorMessage);
        }

        private void btnBaslat_Click(object sender, EventArgs e)
        {
            timer.Start();
            LogMesaji("Excel dosyalarını bekliyorum.");
        }

        private void btnDurdur_Click(object sender, EventArgs e)
        {
            timer.Stop();
            LogMesaji("Kontrol ve taşıma durduruldu.");
        }
    }
}



--



private void AddColoredTextToListBox(string text)
{
    if (InvokeRequired)
    {
        Invoke(new Action(() => AddColoredTextToListBox(text)));
    }
    else
    {
        // Farklı renklerin bulunduğu bir dizi
        Color[] colors = { Color.Red, Color.Green, Color.Blue, Color.Orange, Color.Purple };

        // ListBox'a satır ekle
        listBoxLog.Items.Add(text);

        // Eklenen satırın indeksi
        int index = listBoxLog.Items.Count - 1;

        // Renk indeksi (satır indeksi ile renk dizisi uzunluğunun modu alınarak döngü yapılır)
        int colorIndex = index % colors.Length;

        // Satırın rengini belirle
        listBoxLog.SetSelected(index, true);
        listBoxLog.SelectedIndices.Clear(); // Seçimi temizle
        listBoxLog.SelectedIndices.Add(index); // İlgili satırı seç

        // Satırın rengini değiştir
        listBoxLog.SelectionColor = colors[colorIndex];
    }
}


private void btnBaslat_Click(object sender, EventArgs e)
{
    timer.Start();
    LogColoredMessage("Excel dosyalarını bekliyorum.");
}

private void btnDurdur_Click(object sender, EventArgs e)
{
    timer.Stop();
    LogColoredMessage("Kontrol ve taşıma durduruldu.");
}

private void LogColoredMessage(string message)
{
    AddColoredTextToListBox(message);
}


---------


private void MoveFileWithFileStream(string sourcePath, string destinationPath)
{
    try
    {
        using (FileStream sourceStream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
        {
            using (FileStream destinationStream = new FileStream(destinationPath, FileMode.Create, FileAccess.Write))
            {
                byte[] buffer = new byte[4096];
                int bytesRead;

                while ((bytesRead = sourceStream.Read(buffer, 0, buffer.Length)) > 0)
                {
                    destinationStream.Write(buffer, 0, bytesRead);
                }
            }
        }

        // Dosya başarıyla taşındıysa, orijinal dosyayı sil
        File.Delete(sourcePath);
    }
    catch (Exception ex)
    {
        LogColoredMessage($"Dosya taşıma hatası: {ex.Message}", Color.Red);
    }
}



private void KontrolEtVeTasi()
{
    lock (kuyrukLock)
    {
        if (islemDevamEdiyor)
            return;

        string[] excelDosyalari = Directory.GetFiles(kaynakDizin, "*.xlsx");
        foreach (string excelDosyasi in excelDosyalari)
        {
            dosyaKuyrugu.Enqueue(excelDosyasi);
        }

        islemDevamEdiyor = true;
    }

    while (dosyaKuyrugu.Count > 0)
    {
        string excelDosyasi;
        lock (kuyrukLock)
        {
            excelDosyasi = dosyaKuyrugu.Dequeue();
        }

        string dosyaAdi = Path.GetFileName(excelDosyasi);
        string hedefDosyaYolu = Path.Combine(hedefDizin, dosyaAdi);

        try
        {
            MoveFileWithFileStream(excelDosyasi, hedefDosyaYolu);

            LogColoredMessage($"Dosya taşındı: {dosyaAdi}", Color.Green);

            // Veritabanına ekle
            ImportAndInsertToDatabase(hedefDosyaYolu);
        }
        catch (Exception ex)
        {
            LogColoredMessage($"Hata oluştu: {ex.Message}", Color.Red);
            LogErrorToFile(ex, dosyaAdi);

            // Hata durumunda Excel dosyasını hatalilar dizinine taşı
            MoveToErrorFolder(excelDosyasi);
        }
    }

    lock (kuyrukLock)
    {
        dosyaKuyrugu.Clear();
    }

    islemDevamEdiyor = false;
}

private void LogColoredMessage(string message, Color color)
{
    if (InvokeRequired)
    {
        Invoke(new Action(() => LogColoredMessage(message, color)));
    }
    else
    {
        listBoxLog.SelectionStart = listBoxLog.TextLength;
        listBoxLog.SelectionLength = 0;

        listBoxLog.SelectionColor = color;
        listBoxLog.AppendText(message + Environment.NewLine);

        listBoxLog.SelectionColor = listBoxLog.ForeColor;
    }
}


Animasyonlar için

using System;
using System.Drawing;
using System.Windows.Forms;

public partial class Form1 : Form
{
    private Timer animationTimer = new Timer();
    private int currentFrame = 0;

    public Form1()
    {
        InitializeComponent();

        animationTimer.Interval = 100; // Her çerçeve arası süre
        animationTimer.Tick += AnimationTimer_Tick;
    }

    private void AnimationTimer_Tick(object sender, EventArgs e)
    {
        // Her çerçeve için gerekli işlemleri yapın

        // PictureBox'a çerçeveyi çizin
        pictureBox.Image = GetNextFrame();
    }

    private Bitmap GetNextFrame()
    {
        // Çerçeveyi döndürmek için gerekli işlemleri yapın
        // Örneğin, resimlerinizi bir dizi içinde tutun ve sırayla gösterin

        // Örnek olarak resim dizisi kullanma:
        string[] imagePaths = { "frame1.png", "frame2.png", "frame3.png" };

        Bitmap frame = new Bitmap(imagePaths[currentFrame]);

        // Dizinin sonuna geldiysek sıfıra geri dön
        currentFrame = (currentFrame + 1) % imagePaths.Length;

        return frame;
    }

    private void btnStartAnimation_Click(object sender, EventArgs e)
    {
        animationTimer.Start();
    }

    private void btnStopAnimation_Click(object sender, EventArgs e)
    {
        animationTimer.Stop();
    }
}








