using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;

namespace AutoGenerateErrorMapJson.ErrorMapping
{
    public static class ErrorMapFileBuilder
    {
        public static string BuildErrorMapsJson(string filename, bool isForShell)
        {
            string errorMapsJson = null;

            ResourceManager rm = new ResourceManager(typeof(FaultCodes));
            PropertyInfo[] pis = typeof(FaultCodes).GetProperties(BindingFlags.Public | BindingFlags.Static);
            IEnumerable<KeyValuePair<string, string>> values =
                from pi in pis
                where pi.PropertyType == typeof(string)
                select new KeyValuePair<string, string>(
                    pi.Name,
                    rm.GetString(pi.Name));
            Dictionary<string, string> FaultCodesDictionary = values.ToDictionary(k => k.Key, v => v.Value);

            List<ExcelObject> lstExcelObject = ImportExcel(filename, isForShell);

            StringBuilder errorMapJsonString = new StringBuilder();
            StringWriter stringWriterInstance = new StringWriter(errorMapJsonString);
            using (JsonWriter writer = new JsonTextWriter(stringWriterInstance))
            {


                writer.Formatting = Formatting.Indented;
                writer.WriteStartObject();
                writer.WritePropertyName("SupplierErrorMapModel");
                writer.WriteStartObject();
                int i = 0;
                foreach (var excelObj in lstExcelObject)
                {
                    if (i == 0)
                    {
                        i++;
                        continue;
                    }

                    var supplierError = string.Empty;
                    try
                    {
                        var keyVal = FaultCodesDictionary.Single(f => f.Value == excelObj.ErrorCode);

                    if (!isForShell)
                    {
                        supplierError = excelObj?.SupplierError;
                        writer.WritePropertyName(supplierError); // writing supplier error code as key
                    }
                    else
                    {
                        supplierError = new Random().Next(1, 10).ToString() + keyVal.Value.ToString();
                        writer.WritePropertyName(supplierError); // writing supplier error code as key 
                    }

                    writer.WriteStartObject();
                    writer.WritePropertyName("ErrorCode");
                    writer.WriteValue(keyVal.Value.ToString());
                    writer.WritePropertyName("ErrorResourceName");
                    writer.WriteValue(keyVal.Key.ToString());
                    writer.WritePropertyName("SupplierError");
                    writer.WriteValue(supplierError);
                    writer.WritePropertyName("HttpStatusCode");
                    writer.WriteValue(excelObj?.HttpStatusCode);
                    writer.WritePropertyName("Category");
                    writer.WriteValue(excelObj?.ErrorCategory);
                    writer.WritePropertyName("ErrorType");
                    writer.WriteValue(excelObj?.ErrorType);
                    writer.WriteEndObject();
                    i++;
                    }
                    catch (InvalidOperationException)
                    {
                        i++;
                        continue;
                    }
                }

                writer.WriteEnd();
                writer.WriteEndObject();
                Debug.WriteLine(stringWriterInstance);
                errorMapsJson = stringWriterInstance.ToString();
            }

            return errorMapsJson;
        }

        public static List<ExcelObject> ImportExcel(string filename, bool isForShell)
        {
            string errorMapsExcelFilePath = AppContext.BaseDirectory + "\\ErrorMapping\\" + filename;
            FileInfo errorMapsExcelFileInfo = new FileInfo(errorMapsExcelFilePath);
            List<ExcelObject> listExcelObject = new List<ExcelObject>();
            try
            {
                using (ExcelPackage package = new ExcelPackage(errorMapsExcelFileInfo))
                {
                    StringBuilder excelData = new StringBuilder();
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                    int rowCount = worksheet.Dimension.Rows;
                    int ColCount = worksheet.Dimension.Columns;
                    ColCount = isForShell ? 4 : 5;
                    bool bHeaderRow = true;

                    for (int row = 1; row <= rowCount; row++)
                    {
                        ExcelObject excelObject = new ExcelObject();
                        for (int col = 1; col <= ColCount; col++)
                        {
                            if (bHeaderRow)
                            {
                                switch (col)
                                {
                                    case 1:
                                        excelObject.ErrorCode = worksheet.Cells[row, col].Value.ToString();
                                        break;
                                    case 2:
                                        excelObject.HttpStatusCode = worksheet.Cells[row, col].Value.ToString();
                                        break;
                                    case 3:
                                        excelObject.ErrorCategory = worksheet.Cells[row, col].Value.ToString();
                                        break;
                                    case 4:
                                        excelObject.ErrorType = worksheet.Cells[row, col].Value.ToString();
                                        break;
                                    case 5:
                                        if(!isForShell)
                                            excelObject.SupplierError = worksheet.Cells[row, col].Value.ToString();
                                        break;
                                    default: break;
                                }

                                excelData.Append(worksheet.Cells[row, col].Value.ToString() + "\t");
                            }
                            else
                            {
                                excelData.Append(worksheet.Cells[row, col].Value.ToString() + "\t");
                            }
                        }
                        listExcelObject.Add(excelObject);
                        excelData.Append(Environment.NewLine);
                    }
                }
            }
            catch (Exception exception)
            {
                Debug.WriteLine("Some error occured while importing." + exception.Message);
            }

            return listExcelObject;
        }
    }
}
