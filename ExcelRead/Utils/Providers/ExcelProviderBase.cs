using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelRead.Utils.Providers
{
    public abstract class ExcelProviderBase<T> : IDisposable
    {
        #region Metods

        /// <summary>
        /// Excel dosyasını açar ver yollar.
        /// </summary>
        /// <param name="path">Excel in yolu</param>
        /// <returns>Açılmış excel dosyası.</returns>
        protected ExcelWorksheet GetWorkSheet(string path)
        {
            try
            {
                ExcelPackage excelPackage = new ExcelPackage();
                FileStream fileStream = new FileStream(path, FileMode.Open);
                excelPackage.Load(fileStream);
                ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.First();

                return excelWorksheet;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "HATA", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return null;
        }

        /// <summary>
        /// Okuma işlemi tamamlandığında nesneleri bellekten atar.
        /// </summary>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
            GC.Collect();
        }

        #endregion
    }
}

