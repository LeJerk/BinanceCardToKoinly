using OfficeOpenXml;

namespace BinanceCardToKoinly
{
    internal static class Program
    {
        private const string BINANCE_FILE = "C:\\temp\\binance\\binance_card_transactions.xlsx";
        private const string KOINLY_FILE = "C:\\temp\\binance\\binance_card_koinly_export.xlsx";

        private static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Create excel for Koinly
            using var koinlyExcel = new ExcelPackage();
            var koinlySheet = koinlyExcel.Workbook.Worksheets.Add("sheet1");

            koinlySheet.Cells["A1"].Value = "Koinly Date";
            koinlySheet.Cells["B1"].Value = "Amount";
            koinlySheet.Cells["C1"].Value = "Currency";
            koinlySheet.Cells["D1"].Value = "Description";

            // Read excel from Binance
            using var binanceExcel = new ExcelPackage(new FileInfo(fileName: BINANCE_FILE));
            var binanceSheet = binanceExcel.Workbook.Worksheets["sheet1"];
            var binanceStart = binanceSheet.Dimension.Start;
            var binanceEnd = binanceSheet.Dimension.End;

            for (int row = binanceStart.Row; row <= binanceEnd.Row; row++)
            {
                if (row == 1)
                {
                    continue;
                }

                // Set date
                string[] dateParts = binanceSheet.Cells[row, 1].Text.Split(' ');
                string[] clockParts = dateParts[3].Split(':');

                var transactionDate = new DateTime(
                    year: Convert.ToInt32(dateParts[5]),
                    month: GetMonth(dateParts[1]),
                    day: Convert.ToInt32(dateParts[2]),
                    hour: Convert.ToInt32(clockParts[0]),
                    minute: Convert.ToInt32(clockParts[1]),
                    second: Convert.ToInt32(clockParts[2])
                );

                koinlySheet.Cells[row, 1].Value = $"{transactionDate} UTC";

                // Set amount and currency
                string[] assetUsed = binanceSheet.Cells[row,6].Text.Split(' ');

                koinlySheet.Cells[row, 2].Value = $"-{assetUsed[1]}";
                koinlySheet.Cells[row, 3].Value = assetUsed[0];

                // Set description
                koinlySheet.Cells[row, 4].Value = binanceSheet.Cells[row, 2].Value;
            }

            koinlyExcel.SaveAs(new FileInfo(KOINLY_FILE));
        }

        private static int GetMonth(string month) => month switch
        {
            "Jan" => 1,
            "Feb" => 2,
            "Mar" => 3,
            "Apr" => 4,
            "May" => 5,
            "Jun" => 6,
            "Jul" => 7,
            "Aug" => 8,
            "Sep" => 9,
            "Oct" => 10,
            "Nov" => 11,
            "Dec" => 12,
            _ => 0,
        };
    }
}