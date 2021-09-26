using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;


namespace Kripto_Para_Toplu_Veri_Çekme_V2
{
    class Program
    {
        static void Main(string[] args)
        {
            Random rnd = new Random();
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Title = "Kripto Para Toplu Veri Çekme V2";
            for (; ; )
            {
                Console.WriteLine("Website : kodzamani.weebly.com");
                Console.WriteLine("İnstagram : @kodzamani.tk");
                Console.WriteLine("------------------------------");
                Console.Write("Kripto para :");
                string para = Console.ReadLine();
                Stopwatch süre = new Stopwatch();
                süre.Start();
                Console.WriteLine("------------------------------");
                int sayi = rnd.Next(999999999);
                try
                {
                    Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                    if (xlApp == null)
                    {
                        Console.WriteLine("Bilgisayarınızda excel kurulu değil.");
                        return;
                    }
                    Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                    Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                    object misValue = System.Reflection.Missing.Value;
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    xlWorkSheet.Cells[1, 6] = "Bu veriler 'Kripto Para Toplu Veri Çekme V2' programı ile çekilmiştir website : kodzamani.weebly.com, instagram : @kodzamani.tk";
                    xlWorkSheet.Cells[2, 1] = "Date";
                    xlWorkSheet.Cells[2, 2] = "Market Cap";
                    xlWorkSheet.Cells[2, 3] = "Volume";
                    xlWorkSheet.Cells[2, 4] = "Open";
                    xlWorkSheet.Cells[2, 5] = "Close";
                    using (WebClient client = new WebClient())
                    {
                        client.Encoding = Encoding.UTF8;
                        string html = client.DownloadString("https://www.coingecko.com/en/coins/" + para + "/historical_data/usd?end_date=3000-02-24&start_date=2000-02-24");
                        HtmlAgilityPack.HtmlDocument htmlDocument = new HtmlAgilityPack.HtmlDocument();
                        htmlDocument.LoadHtml(html);
                        HtmlNodeCollection tarihler = htmlDocument.DocumentNode.SelectNodes("//th[@class='font-semibold text-center']");
                        HtmlNodeCollection veriler = htmlDocument.DocumentNode.SelectNodes("//td[@class='text-center']");
                        if (tarihler != null && veriler != null)
                        {
                            for (int a = 3; ; a++)
                            {
                                if (veriler.Count == 0 || tarihler.Count == 0)
                                    break;
                                string date = tarihler[0].InnerText.Replace("\r", "").Replace("\t", "").Replace("\n","");
                                string Market_cap = veriler[0].InnerText.Replace("\r", "").Replace("\t", "").Replace("\n", "");
                                string Volume = veriler[1].InnerText.Replace("\r", "").Replace("\t", "").Replace("\n", "");
                                string Open = veriler[2].InnerText.Replace("\r", "").Replace("\t", "").Replace("\n", "");
                                string Close = veriler[3].InnerText.Replace("\r", "").Replace("\t", "").Replace("\n", "");
                                try
                                {
                                    xlWorkSheet.Cells[a, 1] = date;
                                    xlWorkSheet.Cells[a, 2] = Market_cap;
                                    xlWorkSheet.Cells[a, 3] = Volume;
                                    xlWorkSheet.Cells[a, 4] = Open;
                                    xlWorkSheet.Cells[a, 5] = Close;
                                }
                                catch { }
                                tarihler.RemoveAt(0);
                                for (int i = 0; i < 4; i++)
                                    veriler.RemoveAt(0);
                                Console.WriteLine($"Date :{date} Market Cap :{Market_cap} Volume :{Volume} Open :{Open} Close :{Close}");
                            }
                        }
                    }
                    xlWorkBook.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "kodzamani-" + sayi + ".xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlWorkSheet);
                    Marshal.ReleaseComObject(xlWorkBook);
                    Marshal.ReleaseComObject(xlApp);
                    Console.WriteLine("------------------------------");
                    Console.WriteLine("Excel dosyası başarıyla oluşturuldu.");
                    Process.Start(AppDomain.CurrentDomain.BaseDirectory + "kodzamani-" + sayi + ".xls");
                }
                catch
                {
                    Console.WriteLine("Verileriniz çekilemedi. https://www.coingecko.com/en/coins/ linkine tıklayın. Seçtiğiniz kripto paranın üzerine basın ve coins/ tan sonra yazan kripto para ismini programa yazın.");
                    Console.WriteLine("------------------------------");
                    Console.WriteLine("Excel dosyası oluşturulamadı.");
                }
                süre.Stop();
                Console.WriteLine("Geçen süre :" + süre.Elapsed.TotalSeconds.ToString() + " s");
                Console.WriteLine("Farklı bir işlem yapmak için herhangi bir tuşa basın.");
                Console.ReadLine();
                Console.Clear();
            }
        }
    }
}
