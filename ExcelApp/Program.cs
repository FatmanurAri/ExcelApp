// See https://aka.ms/new-console-template for more information
using OfficeOpenXml;


Console.WriteLine("Hello, World!");

using (var package = new ExcelPackage(new FileInfo("C:\\Users\\Fatmanur_Ari\\Downloads\\202404R2.xlsx")))
{
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    var worksheet = package.Workbook.Worksheets[0];
    var Sayi = 0;
    for (int row = 3; row <= worksheet.Dimension.End.Row; row++)
    {
        Sayi += 1;

        var markaKodu = Convert.ToInt32(worksheet.Cells[row, 1].Value);
        var markaAdi = worksheet.Cells[row, 3].Value.ToString();

        var tipKodu = Convert.ToInt32(worksheet.Cells[row, 2].Value);
        var tipAdi = worksheet.Cells[row, 4].Value.ToString();

        // logger.LogError($"{markaKodu} / {tipKodu}");


        for (int col = 5; col <= worksheet.Dimension.End.Column; col++)
        {
            var yil = Convert.ToInt32(worksheet.Cells[2, col].Value);
            var bedel = Convert.ToDouble(worksheet.Cells[row, col].Value);
            if (bedel == 0) continue;


            // logger.LogError($"{markaAdi} - {tipAdi} {yil} / {Convert.ToInt32(worksheet.Cells[row, col].Value)}");
           Console.WriteLine($"{Sayi} - {markaAdi} - {tipAdi} {yil} / {Convert.ToInt32(worksheet.Cells[row, col].Value)}");
            //logger.LogError($"{KayitSayisi} / {Sayi}");


            //if (kasko == null)
            //{
            //    await fILO_KT_KASKORepository.AddAsync(new FILO_KT_KASKO
            //    {
            //        MarkaKodu = markaKodu,
            //        MarkaAdi = markaAdi,
            //        TipKodu = tipKodu,
            //        TipAdi = tipAdi,
            //        Yil = yil,
            //        Bedel = bedel,
            //        AracKodu = $"{markaKodu.ToString().PadLeft(3, '0')}{tipKodu.ToString().PadLeft(4, '0')}",
            //        GuncellemeTarihi = DateTime.Now,
            //        DosyaAdi = file.FileName,
            //    });
            //}
            //else
            //{
            //    kasko.MarkaAdi = markaAdi;
            //    kasko.TipAdi = tipAdi;
            //    kasko.Bedel = bedel;
            //    kasko.AracKodu = $"{markaKodu.ToString().PadLeft(3, '0')}{tipKodu.ToString().PadLeft(4, '0')}";
            //    kasko.GuncellemeTarihi = DateTime.Now;
            //    kasko.DosyaAdi = file.FileName;
            //    await fILO_KT_KASKORepository.UpdateAsync(kasko);
            //}
        }
    }
}

