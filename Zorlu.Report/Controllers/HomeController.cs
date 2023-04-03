using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.Globalization;
using System.Threading;
using OfficeOpenXml;
using Zorlu.Report.Models;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Hosting;
using System.Numerics;
using Microsoft.EntityFrameworkCore;
using Zorlu.Report.Models.Contexts;

namespace Zorlu.Report.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public const string aktiviteGirisTalepleri = $@" '1258954', '1260500', '1258213', '1249355', '723650', '1249084', '1248940', '1249359', '1255026', '1249336', '1255000', '1249348', '1255021', '1254995', '1249351', '1255004', '1255250', '1249340', '865277', '873815', '873826', '874555', '875270', '878531', '882532', '882533', '882535', '1171994', '1171995', '1173470', '1173474', '1173477', '1173479', '1173481', '1173483', '1176001', '1176005', '1180050', '1195930', '1204511', '1216367', '1218145', '1224461', '1227244', '1243715', '1181678'";

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [HttpGet]
        public async Task<IActionResult> ReportOne(string raporOncekiAy, string raporAyTarihi, string raporBaslangic = "2021-01-01")
        {
            var context = new ZorluContext();

            var ayOncesiAcik = @$"(SELECT COUNT(DISTINCT D.istekId)
FROM Data AS [D] WITH (NOLOCK)
WHERE D.talepProje = 'Y'
AND ISNULL(D.sonSinifUygulama, '') != 'Aktivite Giriş'
AND D.istekTarihi BETWEEN CAST('{raporBaslangic}' AS DATETIME) AND CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('{raporOncekiAy}')), ' 23:59:59'))
AND (d.istekKapanisTarihi IS NULL OR D.istekKapanisTarihi >= CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('{raporAyTarihi}')), ' 23:59:59')))
AND (COALESCE(ISNULL(D.istekKapanisTarihi, NULL), D.istekDurumu) != '0')
AND HASHBYTES('MD5', musteriTemsilcileri) != 0xE258A228A07CD38926067F3495B5FA9D --BoraSelcanoğlu
AND D.istekId NOT IN ({aktiviteGirisTalepleri}))";

            var ayOncesiKapali = $@"(SELECT COUNT(DISTINCT D.istekId)
FROM Data AS [D] WITH (NOLOCK)
WHERE D.talepProje = 'Y'
AND ISNULL(D.sonSinifUygulama, '') != 'Aktivite Giriş'
AND D.istekTarihi BETWEEN CAST('{raporBaslangic}' AS DATE) AND CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('{raporOncekiAy}')), ' 23:59:59'))
AND (D.istekKapanisTarihi BETWEEN DATEADD(mm, DATEDIFF(mm, 0, '{raporAyTarihi}'), 0)  AND  CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('{raporAyTarihi}')), ' 23:59:59')))
AND HASHBYTES('MD5', musteriTemsilcileri) != 0xE258A228A07CD38926067F3495B5FA9D --BoraSelcanoğlu
AND D.istekId NOT IN ({aktiviteGirisTalepleri}))";

            var ayAcik = $@"(SELECT COUNT(DISTINCT D.istekId)
FROM Data AS [D] WITH (NOLOCK)
WHERE D.talepProje = 'Y'
AND ISNULL(D.sonSinifUygulama, '') != 'Aktivite Giriş'
AND D.istekTarihi BETWEEN DATEADD(mm, DATEDIFF(mm, 0, '{raporAyTarihi}'), 0) AND CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('{raporAyTarihi}')), ' 23:59:59'))
AND (d.istekKapanisTarihi IS NULL OR D.istekKapanisTarihi >= CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('{raporAyTarihi}')), ' 23:59:59')))
AND (COALESCE(ISNULL(D.istekKapanisTarihi, NULL), D.istekDurumu) != '0')
AND HASHBYTES('MD5', musteriTemsilcileri) != 0xE258A228A07CD38926067F3495B5FA9D --BoraSelcanoğlu
AND D.istekId NOT IN ({aktiviteGirisTalepleri}))";

            var ayKapali = $@"(SELECT COUNT(DISTINCT D.istekId)
FROM Data AS [D] WITH (NOLOCK)
WHERE D.talepProje = 'Y'
AND ISNULL(D.sonSinifUygulama, '') != 'Aktivite Giriş'
AND D.istekTarihi BETWEEN DATEADD(mm, DATEDIFF(mm, 0, '{raporAyTarihi}'), 0) AND CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('{raporAyTarihi}')), ' 23:59:59'))
AND (D.istekKapanisTarihi BETWEEN DATEADD(mm, DATEDIFF(mm, 0, '{raporAyTarihi}'), 0)  AND  CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('{raporAyTarihi}')), ' 23:59:59')))
AND HASHBYTES('MD5', musteriTemsilcileri) != 0xE258A228A07CD38926067F3495B5FA9D --BoraSelcanoğlu
AND D.istekId NOT IN ({aktiviteGirisTalepleri}))";

            var uygulamaDestekTalepleri = $@"(SELECT COUNT(DISTINCT D.istekId)
FROM Data AS[D] WITH(NOLOCK)
WHERE D.talepProje = 'Y'
AND D.istekTarihi BETWEEN CAST('2022-01-01' AS DATETIME) AND CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('2023-02-01')), ' 23:59:59'))
AND(d.istekKapanisTarihi IS NULL OR D.istekKapanisTarihi BETWEEN DATEADD(mm, DATEDIFF(mm, 0, '{raporAyTarihi}'), 0)  AND  CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('{raporAyTarihi}')), ' 23:59:59')))
AND HASHBYTES('MD5', musteriTemsilcileri) != 0xE258A228A07CD38926067F3495B5FA9D--BoraSelcanoğlu
AND D.bistekTipi = 'Uygulama Destek'
AND D.istekId NOT IN ({aktiviteGirisTalepleri}))";

            var kullaniciIslemleriTalepleri = $@"(SELECT COUNT(DISTINCT D.istekId)
FROM Data AS[D] WITH(NOLOCK)
WHERE D.talepProje = 'Y'
AND D.istekTarihi BETWEEN CAST('2022-01-01' AS DATETIME) AND CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('2023-02-01')), ' 23:59:59'))
AND(d.istekKapanisTarihi IS NULL OR D.istekKapanisTarihi BETWEEN DATEADD(mm, DATEDIFF(mm, 0, '{raporAyTarihi}'), 0)  AND  CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('{raporAyTarihi}')), ' 23:59:59')))
AND HASHBYTES('MD5', musteriTemsilcileri) != 0xE258A228A07CD38926067F3495B5FA9D--BoraSelcanoğlu
AND D.bistekTipi LIKE 'Kullan%'
AND D.istekId NOT IN ({aktiviteGirisTalepleri}))";

            var projeTalepleri = $@"(SELECT COUNT(DISTINCT D.istekId)
FROM Data AS[D] WITH(NOLOCK)
WHERE D.talepProje = 'Y'
AND D.istekTarihi BETWEEN CAST('2022-01-01' AS DATETIME) AND CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('2023-02-01')), ' 23:59:59'))
AND(d.istekKapanisTarihi IS NULL OR D.istekKapanisTarihi BETWEEN DATEADD(mm, DATEDIFF(mm, 0, '{raporAyTarihi}'), 0)  AND  CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('{raporAyTarihi}')), ' 23:59:59')))
AND HASHBYTES('MD5', musteriTemsilcileri) != 0xE258A228A07CD38926067F3495B5FA9D--BoraSelcanoğlu
AND D.bistekTipi = 'Proje'
AND D.istekId NOT IN ({aktiviteGirisTalepleri}))";

            var elektrikTalepleri = $@"(SELECT COUNT(DISTINCT D.istekId)
FROM Data AS[D] WITH(NOLOCK)
WHERE D.talepProje = 'Y'
AND D.istekTarihi BETWEEN CAST('2022-01-01' AS DATETIME) AND CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('2023-02-01')), ' 23:59:59'))
AND(d.istekKapanisTarihi IS NULL OR D.istekKapanisTarihi BETWEEN DATEADD(mm, DATEDIFF(mm, 0, '{raporAyTarihi}'), 0)  AND  CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('{raporAyTarihi}')), ' 23:59:59')))
AND HASHBYTES('MD5', musteriTemsilcileri) != 0xE258A228A07CD38926067F3495B5FA9D--BoraSelcanoğlu
AND D.sektor = 'Elektrik'
AND D.istekId NOT IN ({aktiviteGirisTalepleri}))";

            var gazTalepleri = $@"(SELECT COUNT(DISTINCT D.istekId)
FROM Data AS[D] WITH(NOLOCK)
WHERE D.talepProje = 'Y'
AND D.istekTarihi BETWEEN CAST('2022-01-01' AS DATETIME) AND CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('2023-02-01')), ' 23:59:59'))
AND(d.istekKapanisTarihi IS NULL OR D.istekKapanisTarihi BETWEEN DATEADD(mm, DATEDIFF(mm, 0, '{raporAyTarihi}'), 0)  AND  CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('{raporAyTarihi}')), ' 23:59:59')))
AND HASHBYTES('MD5', musteriTemsilcileri) != 0xE258A228A07CD38926067F3495B5FA9D--BoraSelcanoğlu
AND D.sektor LIKE '%gaz'
AND D.istekId NOT IN ({aktiviteGirisTalepleri}))";

            var ortakTalepler = $@"(SELECT COUNT(DISTINCT D.istekId)
FROM Data AS[D] WITH(NOLOCK)
WHERE D.talepProje = 'Y'
AND D.istekTarihi BETWEEN CAST('2022-01-01' AS DATETIME) AND CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('2023-02-01')), ' 23:59:59'))
AND(d.istekKapanisTarihi IS NULL OR D.istekKapanisTarihi BETWEEN DATEADD(mm, DATEDIFF(mm, 0, '{raporAyTarihi}'), 0)  AND  CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('{raporAyTarihi}')), ' 23:59:59')))
AND HASHBYTES('MD5', musteriTemsilcileri) != 0xE258A228A07CD38926067F3495B5FA9D--BoraSelcanoğlu
AND D.sektor = 'Ortak'
AND D.istekId NOT IN ({aktiviteGirisTalepleri}))";

            var acikTalepler = @$"(SELECT COUNT(DISTINCT D.istekId)
FROM Data AS [D] WITH (NOLOCK)
WHERE D.talepProje = 'Y'
AND ISNULL(D.sonSinifUygulama, '') != 'Aktivite Giriş'
AND D.istekTarihi BETWEEN CAST('{raporBaslangic}' AS DATETIME) AND CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('{raporAyTarihi}')), ' 23:59:59'))
AND (d.istekKapanisTarihi IS NULL OR D.istekKapanisTarihi >= CONVERT(datetime, CONCAT(CONVERT(date, EOMONTH('{raporAyTarihi}')), ' 23:59:59')))
AND (COALESCE(ISNULL(D.istekKapanisTarihi, NULL), D.istekDurumu) != '0')
AND HASHBYTES('MD5', musteriTemsilcileri) != 0xE258A228A07CD38926067F3495B5FA9D --BoraSelcanoğlu
AND D.istekId NOT IN ({aktiviteGirisTalepleri}))";


            var sql = $@"SELECT 
	[IstekDurumu] = 'Açık'
   ,[AyOncesi] = {ayOncesiAcik}
   ,[Ay] = {ayAcik}
   ,[Total] = {ayOncesiAcik} + {ayAcik}
   ,[Total] = {ayOncesiAcik} + {ayAcik}
FROM (
  VALUES (1)
) AS X(a)
UNION ALL
SELECT 
	[IstekDurumu] = 'Kapalı'
   ,[AyOncesi] = {ayOncesiKapali}
   ,[Ay] = {ayKapali}
   ,[Total] = {ayOncesiKapali} + {ayKapali}
FROM (
  VALUES (0)
) AS X(a)";

            var data = await context.Rapor1.FromSqlRaw(sql).ToListAsync();

            var genelToplam = new Rapor1();
            genelToplam.IstekDurumu = "Genel Toplam";
            genelToplam.AyOncesi = data.Sum(x => x.AyOncesi);
            genelToplam.Ay = data.Sum(x => x.Ay);
            genelToplam.Total = genelToplam.AyOncesi + genelToplam.Ay;

            data.Add(genelToplam);


            return Ok(data);
        }

        [HttpPost]
        public async Task<IActionResult> ImportAsync([FromForm] IFormFile file, CancellationToken cancellationToken)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var cultureUS = new CultureInfo("tr-TR");
            await using var stream = new MemoryStream();
            await file.CopyToAsync(stream, cancellationToken);

            using var package = new ExcelPackage(stream);
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            var rowCount = worksheet.Dimension.Rows;

            var list = new List<Data>();
            for (var row = 2; row <= rowCount; row++)
            {
                var data = new Data();

                data.aktiviteAciklamasi = worksheet.Cells[row, 1].Value?.ToString();
                data.aktiviteId = worksheet.Cells[row, 2].Value?.ToString();
                data.aktiviteKaynakTipi = worksheet.Cells[row, 3].Value?.ToString();
                data.aktiviteMaliyeti = worksheet.Cells[row, 4].Value?.ToString();
                data.aktiviteOnaylayanPyIslemiGerceklestiren = worksheet.Cells[row, 5].Value?.ToString();
                data.aktiviteOnaylayanPy = worksheet.Cells[row, 6].Value?.ToString();
                data.aktiviteSonSinifModulDigerAciklama = worksheet.Cells[row, 7].Value?.ToString();
                data.aktiviteSonSinifModul = worksheet.Cells[row, 8].Value?.ToString();
                data.aktiviteSonSinifTalepIcerigi = worksheet.Cells[row, 9].Value?.ToString();
                data.aktiviteSonSinifTransactionDigerAciklama = worksheet.Cells[row, 10].Value?.ToString();
                data.aktiviteSonSinifTransaction = worksheet.Cells[row, 11].Value?.ToString();
                data.aktiviteSonSinifUygulamaDigerAciklama = worksheet.Cells[row, 12].Value?.ToString();
                data.aktiviteSonSinifUygulama = worksheet.Cells[row, 13].Value?.ToString();
                data.aktiviteSuresi = worksheet.Cells[row, 14].Value?.ToString();
                data.aktiviteTarihi = worksheet.Cells[row, 15].Value?.ToString();
                data.aktiviteyiOnaylayanSektorYoneticisiIslemiGerceklestiren = worksheet.Cells[row, 16].Value?.ToString();
                data.aktiviteyiOnaylayanSektorYoneticisi = worksheet.Cells[row, 17].Value?.ToString();
                data.akisiSonGuncelleyenRPT = worksheet.Cells[row, 18].Value?.ToString();
                data.akisiSonGuncelleyenRealUserRPT = worksheet.Cells[row, 19].Value?.ToString();
                data.akisiSonGuncelleyenRealUser = worksheet.Cells[row, 20].Value?.ToString();
                data.akisiSonGuncelleyen = worksheet.Cells[row, 21].Value?.ToString();
                data.bpcProjeKodu = worksheet.Cells[row, 22].Value?.ToString();
                data.syOnayliSure = worksheet.Cells[row, 23].Value?.ToString();
                data.bistekTipi = worksheet.Cells[row, 24].Value?.ToString();
                data.jIstekAcilisAyi = worksheet.Cells[row, 25].Value?.ToString();
                data.cikGunSuresi = worksheet.Cells[row, 26].Value?.ToString();
                data.raporAyiAktiviteSuresi = worksheet.Cells[row, 27].Value?.ToString();
                data.aktiviteSonSinifTalepIcerigi2 = worksheet.Cells[row, 28].Value?.ToString();
                data.raporSecimFiltresi = worksheet.Cells[row, 29].Value?.ToString();
                data.onayliMaliyetYuzdesi = worksheet.Cells[row, 30].Value?.ToString();
                data.maliyetBuyukOlanlar = worksheet.Cells[row, 31].Value?.ToString();
                data.modulUygulama = worksheet.Cells[row, 32].Value?.ToString();
                data.danismanFirmaUcreti = worksheet.Cells[row, 33].Value?.ToString();
                data.dagitimAnahtarFirmasi = worksheet.Cells[row, 34].Value?.ToString();
                data.dagitimAnahtariAdi = worksheet.Cells[row, 35].Value?.ToString();
                data.dagitimAnahtariId = worksheet.Cells[row, 36].Value?.ToString();
                data.dagitimAnahtariKullanilmisMi = worksheet.Cells[row, 37].Value?.ToString();
                data.dagitimAnahtariYuzde = worksheet.Cells[row, 38].Value?.ToString();
                data.gorevAciklamasi = worksheet.Cells[row, 39].Value?.ToString();
                data.gorevBtdNotu = worksheet.Cells[row, 40].Value?.ToString();
                data.gorevBaslangicTarihi = worksheet.Cells[row, 41].Value?.ToString();
                data.gorevBitisTarihi = worksheet.Cells[row, 42].Value?.ToString();
                data.gorevGerceklesenBaslangicTarihi = worksheet.Cells[row, 43].Value?.ToString();
                data.GorevGerceklesenBitisTarihi = worksheet.Cells[row, 44].Value?.ToString();
                data.gorevId = worksheet.Cells[row, 45].Value?.ToString();
                data.gorevOnayliOngorulenSure = worksheet.Cells[row, 46].Value?.ToString();
                data.gorevStatusu = worksheet.Cells[row, 47].Value?.ToString();
                data.gorevTahminiBaslangicTarihi = worksheet.Cells[row, 48].Value?.ToString();
                data.gorevTahminiBitisTarihi = worksheet.Cells[row, 49].Value?.ToString();
                data.gorevTahminiSure = worksheet.Cells[row, 50].Value?.ToString();
                data.gorevToplamGerceklesenSure = worksheet.Cells[row, 51].Value?.ToString();
                data.holdingAktiviteSuresi = worksheet.Cells[row, 52].Value?.ToString();
                data.inOut = worksheet.Cells[row, 53].Value?.ToString();
                data.kaynakAktiviteYuzdesi = worksheet.Cells[row, 54].Value?.ToString();
                data.kaynakAtamasiYapilmisMi = worksheet.Cells[row, 55].Value?.ToString();
                data.kaynakFaturalamaTuru = worksheet.Cells[row, 56].Value?.ToString();
                data.kaynakFirmasiFiltre = worksheet.Cells[row, 57].Value?.ToString();
                data.kaynakFirmasi = worksheet.Cells[row, 58].Value?.ToString();
                data.kaynakUzmanlik = worksheet.Cells[row, 59].Value?.ToString();
                data.kaynak = worksheet.Cells[row, 60].Value?.ToString();
                data.kritiklikSeviyesi = worksheet.Cells[row, 61].Value?.ToString();
                data.masrafYeriAciklama = worksheet.Cells[row, 62].Value?.ToString();
                data.masrafYeri = worksheet.Cells[row, 63].Value?.ToString();
                data.masrafYerleriAciklama = worksheet.Cells[row, 64].Value?.ToString();
                data.masrafYerleri = worksheet.Cells[row, 65].Value?.ToString();
                data.musteriTemsilcileri = worksheet.Cells[row, 66].Value?.ToString();
                data.normalDagilimYuzdesi = worksheet.Cells[row, 67].Value?.ToString();
                data.numberOfRecords = worksheet.Cells[row, 68].Value?.ToString();
                data.onaylanmamisAktiviteMaliyeti = worksheet.Cells[row, 69].Value?.ToString();
                data.onaylanmisAktiviteMaliyeti = worksheet.Cells[row, 70].Value?.ToString();
                data.pyAktiviteOnayTarihi = worksheet.Cells[row, 71].Value?.ToString();
                data.pyAktiviteOnayladiMi = worksheet.Cells[row, 72].Value?.ToString();
                data.pyOnayliAktiviteSuresi = worksheet.Cells[row, 73].Value?.ToString();
                data.paraBirimi = worksheet.Cells[row, 74].Value?.ToString();
                data.paraTipiId = worksheet.Cells[row, 75].Value?.ToString();
                data.planTarih = worksheet.Cells[row, 76].Value?.ToString();
                data.projeAdi = worksheet.Cells[row, 77].Value?.ToString();
                data.projeDurumu = worksheet.Cells[row, 78].Value?.ToString();
                data.projeFaturaTipi = worksheet.Cells[row, 79].Value?.ToString();
                data.projeID = worksheet.Cells[row, 80].Value?.ToString();
                data.projeSponsoru = worksheet.Cells[row, 81].Value?.ToString();
                data.projeYoneticileri = worksheet.Cells[row, 82].Value?.ToString();
                data.sektorYoneticisiAktiviteOnayTarihi = worksheet.Cells[row, 83].Value?.ToString();
                data.sektorYoneticisiAktiviteyiOnayladiMi = worksheet.Cells[row, 84].Value?.ToString();
                data.sektorYoneticisininOnayladigiAktiviteSuresi = worksheet.Cells[row, 85].Value?.ToString();
                data.sektor = worksheet.Cells[row, 85].Value?.ToString();
                data.sonSinirModulDigerAcik = worksheet.Cells[row, 87].Value?.ToString();
                data.sonSinifModul = worksheet.Cells[row, 88].Value?.ToString();
                data.sonSinifTalepIcerigi = worksheet.Cells[row, 89].Value?.ToString();
                data.sonSinifTransactionDigerAciklama = worksheet.Cells[row, 90].Value?.ToString();
                data.sonSinifTransaction = worksheet.Cells[row, 91].Value?.ToString();
                data.sonSinifUygulamaDigerAciklama = worksheet.Cells[row, 92].Value?.ToString();
                data.sonSinifUygulama = worksheet.Cells[row, 93].Value?.ToString();
                data.talepProje = worksheet.Cells[row, 94].Value?.ToString();
                data.taskiSonGuncelleyenRPT = worksheet.Cells[row, 95].Value?.ToString();
                data.taskiSonGuncelleyen = worksheet.Cells[row, 96].Value?.ToString();
                data.toplamOnaylanmamisAktiviteMaliyeti = worksheet.Cells[row, 97].Value?.ToString();
                data.toplamOnaylanmisAktiviteMaliyeti = worksheet.Cells[row, 98].Value?.ToString();
                data.onSinifModulDigerAciklama = worksheet.Cells[row, 99].Value?.ToString();
                data.onSinifModul = worksheet.Cells[row, 100].Value?.ToString();
                data.onSinifTalepIcerigi = worksheet.Cells[row, 101].Value?.ToString();
                data.onSinifTransactionDiger = worksheet.Cells[row, 102].Value?.ToString();
                data.onSinifTransaction = worksheet.Cells[row, 103].Value?.ToString();
                data.onSinifUygulamaDigerAciklama = worksheet.Cells[row, 104].Value?.ToString();
                data.onSinifUygulama = worksheet.Cells[row, 105].Value?.ToString();
                data.istekAcanKullanici = worksheet.Cells[row, 106].Value?.ToString();
                data.istekDurumu = worksheet.Cells[row, 107].Value?.ToString();
                data.istekIdRPT = worksheet.Cells[row, 108].Value?.ToString();
                data.istekId = worksheet.Cells[row, 109].Value?.ToString();
                data.istekKapanisTarihi = worksheet.Cells[row, 110].Value?.ToString();
                data.istekKaynakAtamaTarihi = worksheet.Cells[row, 111].Value?.ToString();
                data.istekKaynakYoneticileri = worksheet.Cells[row, 112].Value?.ToString();
                data.istekStatusu = worksheet.Cells[row, 113].Value?.ToString();
                data.istekTanimi = worksheet.Cells[row, 114].Value?.ToString();
                data.istekTarihi = worksheet.Cells[row, 115].Value?.ToString();
                data.istekTerminTarihi = worksheet.Cells[row, 116].Value?.ToString();
                data.istekTipi2 = worksheet.Cells[row, 117].Value?.ToString();
                data.istekWorkflowGuid = worksheet.Cells[row, 118].Value?.ToString();
                data.istekOncelik = worksheet.Cells[row, 119].Value?.ToString();
                data.istekGorevDokumani = worksheet.Cells[row, 120].Value?.ToString();
                data.istegiAcaninYoneticisiId = worksheet.Cells[row, 121].Value?.ToString();
                data.istegiAcaninYoneticisi = worksheet.Cells[row, 122].Value?.ToString();
                data.sirketGroup = worksheet.Cells[row, 123].Value?.ToString();
                data.sirket = worksheet.Cells[row, 124].Value?.ToString();

                list.Add(data);
            }



            var context = new ZorluContext();
            var allData = context.Database.ExecuteSqlRaw($"DELETE FROM Data");

            var datas = list.Select((x, i) => new { Index = i, Value = x })
                .GroupBy(x => x.Index / 1000)
                .Select(x => x.Select(v => v.Value).ToList())
                .ToList();

            try
            {
                foreach (var item in datas)
                {
                    await context.Data.BulkInsertAsync(item, cancellationToken);
                }

            }
            catch (Exception e)
            {
                var msg = e;
            }
            //foreach (var item in list)
            //{

            //}

            return Ok();
            //return result.Success ? (IActionResult)Ok(result) : BadRequest(result.Message);
        }
    }

}