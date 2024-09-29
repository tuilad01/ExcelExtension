using OfficeOpenXml;
using OfficeOpenXmlExtension;

namespace UnitTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void Test_SheetStart1_1WithHeader()
        {
            var filePath = "sheet_start_1_1.xlsx";
            var json = string.Empty;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = package.Workbook.Worksheets[0];

                json = sheet.GetTableJson();
            }

            Assert.AreNotEqual(json, string.Empty);            
        }

        [TestMethod]
        public void Test_SheetStart10_10WithHeader()
        {
            var filePath = "sheet_start_10_10.xlsx";
            var json = string.Empty;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = package.Workbook.Worksheets[0];

                json = sheet.GetTableJson(startRow: 10, startColumn: 10);
            }

            Assert.AreNotEqual(json, string.Empty);
        }

        [TestMethod]
        public void Test_SheetStart10_10WithHeader_Get10Rows()
        {
            var filePath = "sheet_start_10_10.xlsx";
            var json = string.Empty;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = package.Workbook.Worksheets[0];

                json = sheet.GetTableJson(startRow: 10, startColumn: 10, endRow: 19);
            }

            Assert.AreNotEqual(json, string.Empty);
        }

        [TestMethod]
        public void Test_SheetStart10_10WithoutHeader()
        {
            var filePath = "sheet_start_10_10_without_header.xlsx";
            var json = string.Empty;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = package.Workbook.Worksheets[0];

                json = sheet.GetTableJson(startRow:10, startColumn: 10, hasHeader: false);
            }

            Assert.AreNotEqual(json, string.Empty);

        }

        [TestMethod]
        public void Test_SheetStart10_10WithoutHeader_Get10Rows()
        {
            var filePath = "sheet_start_10_10_without_header.xlsx";
            var json = string.Empty;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = package.Workbook.Worksheets[0];

                json = sheet.GetTableJson(startRow: 10, startColumn: 10, endRow: 19, hasHeader: false);
            }

            Assert.AreNotEqual(json, string.Empty);

        }
    }
}