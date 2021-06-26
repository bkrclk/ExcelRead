using ExcelRead.Models;
using ExcelRead.Utils;
using ExcelRead.Utils.Providers;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;

namespace ExcelRead.ViewModel
{
    public class MainViewModel: ViewModelBase
    {
        private ObservableCollection<ExcelReadModel> excelList;
        public ObservableCollection<ExcelReadModel> ExcelList
        {
            get { return excelList; }
            set
            {
                excelList = value;
                NotifyPropertyChanged(nameof(ExcelList));
            }
        }

        private ICommand readExcelCommand;
        private ICommand exportExcelCommand;

        public ICommand ReadExcelCommand
        {
            get
            {
                if (readExcelCommand == null)
                    readExcelCommand = new RelayCommand(ReadExcel);
                return readExcelCommand;
            }
        }


        public ICommand ExportExcelCommand
        {
            get
            {
                if (exportExcelCommand == null)
                    exportExcelCommand = new RelayCommand(ExportExcel);
                return exportExcelCommand;
            }
        }

        private void ReadExcel()
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = @"Excel Dosyaları(*xlsx)|*.xlsx";
            openFile.DefaultExt = ".xlsx";
            openFile.ShowDialog();

            if (string.IsNullOrEmpty(openFile.FileName))
                return;

            ExcelList = new ObservableCollection<ExcelReadModel>(new ExcelProvider().ReadList(openFile.FileName));
        }

        private void ExportExcel()
        {
            try
            {
                MemoryStream memory = new MemoryStream();
                var filePath = Path.Combine(Environment.CurrentDirectory + "\\docs", @"Excelread.xlsx");

                FileStream fileStream = File.OpenRead(filePath);
                fileStream.CopyTo(memory);
                byte[] byteData = memory.ToArray();

                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.FileName = "ExportExcel";
                saveFile.Filter = @"Excel Dosyaları(*xlsx)|*.xlsx";
                saveFile.DefaultExt = ".xslx";
                var dialogResult = saveFile.ShowDialog();
                if (dialogResult == DialogResult.OK)
                {
                    File.WriteAllBytes(saveFile.FileName, byteData);
                    System.Windows.MessageBox.Show("İndirme İşlemi Başarılı.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    System.Windows.MessageBox.Show("İşlem iptal edildi.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("Dosya bulunamadı.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

    }
}

