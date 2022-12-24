using System;
using System.IO;
using System.Linq;
using System.Net.Mail;
using OfficeOpenXml;

namespace ExcelSayfalariAyir
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //Dosyanın adını değiştirin
            string dosyaAdi = "deneme.xlsx";
            //Dosyayı okuyun
            FileInfo dosya = new FileInfo(dosyaAdi);
            using (ExcelPackage excel = new ExcelPackage(dosya))
            {
                //Dosyada bulunan sayfaları dolaşın
                foreach (ExcelWorksheet ws in excel.Workbook.Worksheets)
                {
                    //Sayfanın adını alın
                    string sayfaAdi = ws.Name;
                    //Yeni bir dosya oluşturun ve sayfayı bu dosyaya kaydedin
                    using (ExcelPackage yeniExcel = new ExcelPackage())
                    {
                        yeniExcel.Workbook.Worksheets.Add(sayfaAdi, ws);
                        yeniExcel.SaveAs(new FileInfo(sayfaAdi + ".xlsx"));
                    }
                }
            }
        }
    }
}


//using System;
//using System.IO;
//using System.Linq;
//using OfficeOpenXml;
//using System.Net.Mail;

//namespace ExcelSayfalariAyirVeGonder
//{
//    class Program
//    {
//        static void Main(string[] args)
//        {
//            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
//            //Dosyanın adını değiştirin
//            string dosyaAdi = "deneme.xlsx";
//            //Dosyayı okuyun
//            FileInfo dosya = new FileInfo(dosyaAdi);
//            using (ExcelPackage excel = new ExcelPackage(dosya))
//            {
//                //Dosyada bulunan sayfaları dolaşın
//                foreach (ExcelWorksheet ws in excel.Workbook.Worksheets)
//                {
//                    //Sayfanın adını alın
//                    string sayfaAdi = ws.Name;
//                    //Yeni bir dosya oluşturun ve sayfayı bu dosyaya kaydedin
//                    using (ExcelPackage yeniExcel = new ExcelPackage())
//                    {
//                        yeniExcel.Workbook.Worksheets.Add(sayfaAdi, ws);
//                        yeniExcel.SaveAs(new FileInfo(sayfaAdi + ".xlsx"));
//                    }
//                    //E-posta ayarlarını yapın
//                    MailMessage ePosta = new MailMessage();
//                    ePosta.From = new MailAddress("kendi_mail_adresiniz@example.com");
//                    //Sayfa adını kullanarak e-posta adresini belirleyin
//                    ePosta.To.Add(sayfaAdi + "@example.com");
//                    ePosta.Subject = "Excel Dosyası";
//                    ePosta.Body = "Aşağıdaki dosya sizin için oluşturulmuştur.";
//                    //Ek dosyayı e-postaya ekleyin
//                    Attachment ekDosya = new Attachment(sayfaAdi + ".xlsx");
//                    ePosta.Attachments.Add(ekDosya);
//                    //SMTP sunucusu ayarlarını yapın
//                    SmtpClient smtp = new SmtpClient();
//                    smtp.Host = "smtp.example.com";
//                    smtp.Port = 587;
//                    smtp.Credentials = new System.Net.NetworkCredential("kendi_mail_adresiniz@example.com", "parolanız");
//                    smtp.EnableSsl = true;
//                    //E-postayı gönderin
//                    smtp.Send(ePosta);
//                }
//            }
//        }
//    }
//}