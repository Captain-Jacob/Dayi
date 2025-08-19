using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Windows.Forms;
using ClosedXML.Excel;
using Microsoft.Win32;
using System.Drawing.Printing;
using System.Management;
using System.Runtime.InteropServices;

namespace ProgramKontrol                  //Şimdiden söylüyorum , kodu kontrol edecekler kolay gelsin.
                                           //Bu kod 25.07.2025 de Yakup Nail Ceylan tarafında yapılmıştır
                                            //Günümüzde hata ya da herhangi bir durumda sorumluluk kabul etmemekteyim. 
{
    
        public partial class Form1 : Form
    {
        private Dictionary<string, bool> kontrolSonuclari = new();
        private double yuzdelikOran = 0;
       
        public Form1()
        {
            InitializeComponent();

            BilgileriGoster();
            ProgramlariKontrolEt();

            string kayitDosyasi =  @"\\şirketbilgisi\gizli\KULLANICI_SISTEMLERI\RAPORLAR\Kurulum Kontrol\Kayitlar\ProgramRaporu.xlsx";
                                                                         //bu kayıt dosya yolu ama ana yolumuz değil    
            if (File.Exists(kayitDosyasi))                             //temizlik için ayrı bir yoldan bakıyor 
            {                                                             // uyumlu olması için aynı şekilde güncelleyin
                                                                          //kendi uzantınız için \Kayitlara kadar olan kısmı silin
                TemizleAyniPCKayitlari(kayitDosyasi);
            }
        }

        private void BilgileriGoster()
        {
            string pcAdi = Environment.MachineName;
            string kullaniciAdi = Environment.UserName;
            string windowsSurum = GetWindowsEdition();
            labelBilgi.Text = $"Bilgisayar: {pcAdi} | Kullanıcı: {kullaniciAdi} | Sürüm: {windowsSurum}";
        }
        

        private void ProgramlariKontrolEt()
        {
            var programlar = new List<string>   //Buraya yeni bir dosya eklenecekse o eklenebilir , sorun yok
                                                //okunabilir olsun diye alt alta koydum
                                                    // kontrol edeceğe şeylere buradan bakarsınız
            {
                "Forti Client",
                "Forti Nac",
                "ManageEngine UEMS",
                "Microsoft 365",
                "Microsoft Office",
                "Teams",
                "SAP GUI",
                "Trend Micro",
                "Winrar",
                "Adobe Acrobat",
                "Chrome",
                "LightShot",
                
            };

            kontrolSonuclari = ProgramKontrol.ProgramListesiKontrol(programlar);

            var programGruplari = new List<List<string>>
            {
                new() { "Microsoft 365", "Microsoft Office" }
            };


            foreach (var grup in programGruplari)
            {
                if (grup.Any(p => kontrolSonuclari.ContainsKey(p) && kontrolSonuclari[p]))
                {
                    foreach (var p in grup)
                        kontrolSonuclari[p] = true;
                }
            }



            listBoxYuklu.Items.Clear();
            listBoxEksik.Items.Clear();

            int yuklu = kontrolSonuclari.Count(x => x.Value);
            yuzdelikOran = (double)yuklu / ( programlar.Count + 1 )  * 100; // +1 in nedeni office 1 tane varsa 2.sine gerek yok diye
                                                                           //ignore atması için ama gene buglı ondan 92 gelirse şaşırmayın
            foreach (var item in kontrolSonuclari)
            {
                if (item.Value)
                    listBoxYuklu.Items.Add("✔ " + item.Key);
                else
                    listBoxEksik.Items.Add("✖ " + item.Key);
            }

            labelYuzde.Text = $"{yuzdelikOran:F0}%";        //yüzdeliğin rengi ve oranı
            labelYuzde.BackColor = yuzdelikOran == 100 ? Color.LightGreen :
                                   yuzdelikOran >= 70 ? Color.LightYellow :
                                   Color.DarkRed;
        }

        private string GetWindowsEdition()
        {
            string edition = "Bilinmiyor";
            try
            {
                using (RegistryKey? key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion"))
                {
                    if (key != null)
                    {
                        string? productName = key.GetValue("ProductName")?.ToString();
                        string? build = key.GetValue("CurrentBuild")?.ToString();

                        if (!string.IsNullOrEmpty(productName) && int.TryParse(build, out int buildNumber))
                        {
                            if (buildNumber >= 22000 && !productName.Contains("11"))
                                edition = productName.Replace("10", "11");
                            else
                                edition = productName;
                        }
                        else if (!string.IsNullOrEmpty(productName))
                        {
                            edition = productName;
                        }
                    }
                }
            }
            catch { }

            return edition;
        }

        private void BtnKaydet_Click(object sender, EventArgs e)
        {
            if (kontrolSonuclari == null || kontrolSonuclari.Count == 0)
            {
                MessageBox.Show("Program kontrolü yapılmadı.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (var loading = new LoadingForm())
            {
                loading.Show();
                loading.Refresh();

                KaydetExcel();

                loading.Close();
            }
        }

        public bool WindowsUpdateBekliyorMu()   // Uyarı var mı var , çalışıyo mu ? evet ise dokunma
        {
            try
            {
                Type updateSessionType = Type.GetTypeFromProgID("Microsoft.Update.Session", false);
                if (updateSessionType == null) return false;

                dynamic updateSession = Activator.CreateInstance(updateSessionType);
                dynamic updateSearcher = updateSession.CreateUpdateSearcher();

                dynamic searchResult = updateSearcher.Search("IsInstalled=0 and Type='Software'");
                int count = searchResult.Updates.Count;

                return count > 0; // Bekleyen güncelleme varsa 
            }
            catch
            {
                return false; // Erişim sorunu varsa bekleyen sayılmaz
            }
        }



        private bool ExcelAcilikKontrol(string excelDosyaYolu, out string kullanici, out TimeSpan acikSure)
        {
            kullanici = "Bilinmiyor";
            acikSure = TimeSpan.Zero;

            string dizin = Path.GetDirectoryName(excelDosyaYolu) ?? "";
            string dosyaAdi = Path.GetFileName(excelDosyaYolu);
            string tempDosya = Path.Combine(dizin, "~$" + dosyaAdi);  //bu dosyayı gizli yapmak için 

            if (!File.Exists(tempDosya))
                return false;

            try
            {
                var info = new FileInfo(tempDosya);
                acikSure = DateTime.Now - info.CreationTime;

                var satirlar = File.ReadAllLines(tempDosya);
                if (satirlar.Length > 0 && !string.IsNullOrWhiteSpace(satirlar[0]))
                    kullanici = satirlar[0];
            }
            catch { }

            return true;
        }

        private void TemizleAyniPCKayitlari(string dosyaYolu)
        {
            if (!File.Exists(dosyaYolu)) return;

            using var workbook = new XLWorkbook(dosyaYolu);
            var sheet = workbook.Worksheet(1);

            var satirlar = sheet.RowsUsed().Skip(1)
                .Select(row => new
                {
                    Satir = row,
                    Tarih = DateTime.TryParse(row.Cell(1).GetString(), out var dt) ? dt : DateTime.MinValue,
                    BilgisayarAdi = row.Cell(2).GetString()
                })
                .ToList();

            var gruplar = satirlar.GroupBy(x => x.BilgisayarAdi);

            foreach (var grup in gruplar)
            {
                var enYeni = grup.OrderByDescending(x => x.Tarih).First();

                foreach (var silinecek in grup.Where(x => x != enYeni))
                {
                    silinecek.Satir.Delete();
                }
            }

            workbook.SaveAs(dosyaYolu);
        }

        private int TemizleAyniPCKayitlariVeBosSatiraHazirla(IXLWorksheet sheet, string bilgisayarAdi)
        {
            var satirlar = sheet.RowsUsed().Skip(1)
                .Select(row => new
                {
                    Satir = row,
                    SatirNo = row.RowNumber(),
                    Tarih = DateTime.TryParse(row.Cell(1).GetString(), out var dt) ? dt : DateTime.MinValue,
                    BilgisayarAdi = row.Cell(2).GetString()
                })
                .ToList();

            // Aynı bilgisayar adına sahip satırlar
            var ayniAdliSatirlar = satirlar.Where(x => x.BilgisayarAdi == bilgisayarAdi).ToList();

            // Aynı adda en güncel dışında alayını sil
            if (ayniAdliSatirlar.Count > 1)
            {
                var enYeni = ayniAdliSatirlar.OrderByDescending(x => x.Tarih).First();
                foreach (var s in ayniAdliSatirlar)
                {
                    if (s != enYeni)
                        s.Satir.Delete();
                }
            }

            // Satırı hafif düzenliyor
            if (ayniAdliSatirlar.Count >= 1)
            {
                return ayniAdliSatirlar.OrderByDescending(x => x.Tarih).First().SatirNo;
            }

            // Yeni boş satır bulyo
            int i = 2;
            while (!sheet.Cell(i, 1).IsEmpty())
                i++;

            return i;
        }


        private void KaydetExcel()       // dosyaynın kaydetme yolu
        {
            string klasorYolu = @"\\şirketbilgisi\gizli\KULLANICI_SISTEMLERI\RAPORLAR\Kurulum Kontrol\Kayitlar\ProgramRaporu.xlsx"; //kendi uzantınız için \Kayitlara kadar olan kısmı silin
            string anaYol = Path.Combine(klasorYolu, "ProgramRaporu.xlsx");
            string yedekYol = Path.Combine(klasorYolu, "ProgramRaporu_Yedek.xlsx");

            // Klasör yoksa oluştur
            if (!Directory.Exists(klasorYolu))
                Directory.CreateDirectory(klasorYolu);

            // Ana dosya açıklık kontrolü
            bool anaDosyaAcik = ExcelAcilikKontrol(anaYol, out _, out _);
            bool guncellemeVar = WindowsUpdateBekliyorMu();
            // Eğer ana dosya açıksa , yedek dosya yazılacak
            if (anaDosyaAcik)
            {
                MessageBox.Show("Dayı:Ana dosya kullanımda yeğen, veri yedek dosyaya yazdık.", "Dayı yedeği yaptı!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                TemizleAyniPCKayitlari(yedekYol);
                YazExcelDosyasina(yedekYol, gizliDosyadan: true);
                DosyayiGizle(yedekYol);
                return;
            }

            // Ana dosya açık değilse
            // ve yedek dosya varsa önce onu aktar
            if (File.Exists(yedekYol))
            {
                AktarYedekVerileri(anaYol, yedekYol);
            }

            // Ana dosyada temizlik yap ve datayı yaz
            TemizleAyniPCKayitlari(anaYol);
            YazExcelDosyasina(anaYol);
            MessageBox.Show("Dayı: Çay demlendi yeğen.", "✔ Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }



        private void DosyayiGizle(string yol)
        {
            if (File.Exists(yol))
                File.SetAttributes(yol, FileAttributes.Hidden);
        }

        public static string GetBiosSerialNumber()
        {
            try
            {
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BIOS"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        return obj["SerialNumber"]?.ToString() ?? "Yok";
                    }
                }
            }
            catch
            {
                return "Hata";
            }

            return "Bilinmiyor";
        }


        private void YazExcelDosyasina(string dosyaYolu, bool gizliDosyadan = false)
        {
            bool dosyaVar = File.Exists(dosyaYolu);

            using var workbook = dosyaVar ? new XLWorkbook(dosyaYolu) : new XLWorkbook();
            var sheet = dosyaVar ? workbook.Worksheet(1) : workbook.Worksheets.Add("Kayıtlar");

            string bilgisayarAdi = Environment.MachineName;
            string domainVeyaWorkgroup;



            try
            {
                var envDomain = Environment.UserDomainName;

                // Altaki bir tık kontrolcü gibi bir şey , zaten anlarsınız
                if (string.IsNullOrWhiteSpace(envDomain) || envDomain.Equals(bilgisayarAdi, StringComparison.OrdinalIgnoreCase))
                    domainVeyaWorkgroup = "WORKGROUP";
                else
                    domainVeyaWorkgroup = envDomain;
            }
            catch
            {
                domainVeyaWorkgroup = "Bilinmiyor";
            }

            string yaziciAdi = "Yok";

            try
            {
                var yazicilar = PrinterSettings.InstalledPrinters;
                if (yazicilar.Count > 0)
                    yaziciAdi = string.Join(", ", yazicilar); // Tüm yazıcı adlarını yazması gerek "," ile ayıracak
            }
            catch
            {
                yaziciAdi = "Hata";
            }

            bool tumSuruculerYeni = true;

            try
            {
                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPSignedDriver");
                foreach (ManagementObject obj in searcher.Get())
                {
                    string deviceName = obj["DeviceName"]?.ToString() ?? "";
                    string driverDateStr = obj["DriverDate"]?.ToString() ?? "";

                    if (DateTime.TryParseExact(driverDateStr.Substring(0, 8), "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out DateTime driverDate))
                    {
                        // eğer driver 3 AY güncel değilse her halde yanlıştır , normalde kaç yılda bir update geliyo ki ?
                        if (driverDate < DateTime.Now.AddMonths(-3))
                        {
                            tumSuruculerYeni = false;
                            break;
                        }
                    }
                }
            }
            catch
            {
                tumSuruculerYeni = false; // Hata varsa eski varsayalım
            }


            string tamIsim = $"{bilgisayarAdi}.{domainVeyaWorkgroup}";

            if (!dosyaVar)
            {
                sheet.Cell(1, 1).Value = "Tarih";
                sheet.Cell(1, 2).Value = "Bilgisayar Adı";
                sheet.Cell(1, 3).Value = "Kullanıcı Adı";
                sheet.Cell(1, 4).Value = "IP Adresi";
                sheet.Cell(1, 5).Value = "MAC Adresi";
                sheet.Cell(1, 6).Value = "Yüklü Programlar";
                sheet.Cell(1, 7).Value = "Eksik Programlar";
                sheet.Cell(1, 8).Value = "Windows Sürümü";
                sheet.Cell(1, 9).Value = "Domain Adı";
                sheet.Cell(1, 10).Value = "Yazıcı";
                sheet.Cell(1, 11).Value = "Sürücü Durumu";
                sheet.Cell(1, 12).Value = "Windows Güncellemi";
                sheet.Cell(1, 13).Value = "Seri No";
                sheet.Cell(1, 14).Value = "Yüzdelik Başarı";
                sheet.Cell(1, 15).Value = "Eski Pc";
                sheet.Cell(1, 16).Value = "Eski Seri No";

            }

            string tarihSaat = DateTime.Now.ToString("dd.MM.yyyy HH:mm");
            string pcAdi = Environment.MachineName;
            string kullaniciAdi = Environment.UserName;

            string ip = "Bilinmiyor";
            try
            {
                ip = Dns.GetHostEntry(Dns.GetHostName()).AddressList
                    .FirstOrDefault(x => x.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)?.ToString() ?? "Bilinmiyor";
            }
            catch { }

            string mac = NetworkInterface.GetAllNetworkInterfaces()
                .Where(nic => nic.OperationalStatus == OperationalStatus.Up && nic.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                .Select(nic => string.Join("-", nic.GetPhysicalAddress().GetAddressBytes().Select(b => b.ToString("X2"))))
                .FirstOrDefault() ?? "Bilinmiyor";

            string yuklu = string.Join(", ", kontrolSonuclari.Where(x => x.Value).Select(x => x.Key));
            string eksik = string.Join(", ", kontrolSonuclari.Where(x => !x.Value).Select(x => x.Key));

            // Aynı PC adındaki eski kayıtları sil
            var satirlar = sheet.RowsUsed().Skip(1)
                .Select(row => new
                {
                    Satir = row,
                    SatirNo = row.RowNumber(),
                    Tarih = DateTime.TryParse(row.Cell(1).GetString(), out var dt) ? dt : DateTime.MinValue,
                    BilgisayarAdi = row.Cell(2).GetString()
                })
                .Where(x => x.BilgisayarAdi == pcAdi)
                .ToList();

            if (satirlar.Count > 1)
            {
                var enYeni = satirlar.OrderByDescending(x => x.Tarih).First();
                foreach (var s in satirlar)
                {
                    if (s != enYeni)
                        s.Satir.Delete();
                }
            }

            // Güncel kayıt varsa onun üstüne yaz
            int hedefSatir;
            if (satirlar.Any())
                hedefSatir = satirlar.OrderByDescending(x => x.Tarih).First().SatirNo;
            else
            {
                hedefSatir = 2;
                while (!sheet.Cell(hedefSatir, 1).IsEmpty()) hedefSatir++;
            }

            bool guncellemeVar = WindowsUpdateBekliyorMu();
            string serial = GetBiosSerialNumber();

            // Datayı yaz
            sheet.Cell(hedefSatir, 1).Value = tarihSaat;
            sheet.Cell(hedefSatir, 2).Value = pcAdi;
            sheet.Cell(hedefSatir, 3).Value = kullaniciAdi;
            sheet.Cell(hedefSatir, 4).Value = ip;
            sheet.Cell(hedefSatir, 5).Value = mac;
            sheet.Cell(hedefSatir, 6).Value = yuklu;
            sheet.Cell(hedefSatir, 6).Style.Font.FontColor = XLColor.DarkGreen;
            sheet.Cell(hedefSatir, 7).Value = eksik;
            sheet.Cell(hedefSatir, 7).Style.Font.FontColor = XLColor.Red;
            sheet.Cell(hedefSatir, 8).Value = GetWindowsEdition();
            sheet.Cell(hedefSatir, 9).Value = domainVeyaWorkgroup;
            sheet.Cell(hedefSatir, 10).Value = yaziciAdi;
            sheet.Cell(hedefSatir, 11).Value = tumSuruculerYeni ? "✔" : "✖";
            sheet.Cell(hedefSatir, 11).Style.Font.FontColor = tumSuruculerYeni ? XLColor.DarkGreen : XLColor.Red;
            sheet.Cell(hedefSatir, 12).Value = guncellemeVar ? "✖" : "✔";
            sheet.Cell(hedefSatir, 12).Style.Font.FontColor = guncellemeVar ? XLColor.Red : XLColor.DarkGreen;
            sheet.Cell(hedefSatir, 13).Value = serial;
            sheet.Cell(hedefSatir, 14).Value = $"{yuzdelikOran:F0}%";
            sheet.Cell(hedefSatir, 14).Style.Font.FontColor = XLColor.Blue;
            sheet.Cell(hedefSatir, 15).Value = textBoxNot.Text;
            sheet.Cell(hedefSatir, 16).Value = textBoxNot2.Text;

            try
            {
                workbook.SaveAs(dosyaYolu);
            }
            catch (Exception ex)
            {
                MessageBox.Show("HATA: " + ex.Message);
            }

        }
        
        private void AktarYedekVerileri(string anaYol, string yedekYol)
        {
            using var anaWB = File.Exists(anaYol) ? new XLWorkbook(anaYol) : new XLWorkbook();
            using var yedekWB = new XLWorkbook(yedekYol);

            var anaSheet = anaWB.Worksheets.FirstOrDefault() ?? anaWB.Worksheets.Add("Kayıtlar");
            var yedekSheet = yedekWB.Worksheets.FirstOrDefault();

            if (yedekSheet == null) return;

            int anaSonSatir = anaSheet.LastRowUsed()?.RowNumber() ?? 1;
            int yedekSonSatir = yedekSheet.LastRowUsed()?.RowNumber() ?? 1;

            for (int i = 2; i <= yedekSonSatir; i++)
            {
                for (int j = 1; j <= 17; j++)
                {
                    anaSheet.Cell(anaSonSatir + (i - 1), j).Value = yedekSheet.Cell(i, j).Value;
                }
            }

            anaWB.SaveAs(anaYol);
            File.Delete(yedekYol);
        }
    }
}
//dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
//exe kodu üsteki
//buraya kadar okuduysanız umarım anlamışsınızdır.