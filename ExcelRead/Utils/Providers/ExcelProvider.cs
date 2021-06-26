using ExcelRead.Common;
using ExcelRead.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace ExcelRead.Utils.Providers
{
    public class ExcelProvider : ExcelProviderBase<ExcelReadModel>
    {
        public List<ExcelReadModel> ReadList(string path)
        {
            ExcelWorksheet reader = GetWorkSheet(path);
            List<ExcelReadModel> resultList = new List<ExcelReadModel>();

            if (reader != null)
            {
                for (int i = 2; i <= reader.Dimension.End.Row; i++)
                {
                    if (!string.IsNullOrEmpty(reader.Cells[i, 1].Value.Tostring()))
                    {
                        ExcelReadModel excelReadModel = new ExcelReadModel();

                        try
                        {
                            excelReadModel.TCNo = reader.Cells[i, 1].Value.ToString();
                            excelReadModel.Gain = reader.Cells[i, 2].Value?.ToString();
                            excelReadModel.EducationCode = reader.Cells[i, 3].Value?.ToString();
                            excelReadModel.ProfessionCode = reader.Cells[i, 4].Value?.ToString();

                            resultList.Add(excelReadModel);
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
            }
            base.Dispose();
            return resultList;
        }
    }
}
