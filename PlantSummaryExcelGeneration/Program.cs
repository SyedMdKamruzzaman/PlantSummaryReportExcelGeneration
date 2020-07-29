using System;
using System.IO;


namespace PlantSummaryExcelGeneration
{
    class Program
    {
        public static void Main(string[] args)
        {
            GenerateExcel generateExcel = new GenerateExcel();

            FileStream fileStream = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "Plant Summary Report.xlsx", FileMode.Create, FileAccess.Write);
            MemoryStream memoryStream = generateExcel.GetExcelMemoryStream();
            memoryStream.WriteTo(fileStream);
            fileStream.Close();
            memoryStream.Close();
        } 
    }
}
