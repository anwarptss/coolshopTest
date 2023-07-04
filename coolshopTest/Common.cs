using IronXL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace coolshopTest
{
    internal class Common
    {
        public static DataTable ReadExcel(string fileName)
        {                   

            WorkBook workbook = WorkBook.Load(fileName);
            // Work with a single WorkSheet.
            //you can pass static sheet name like Sheet1 to get that sheet
            //WorkSheet sheet = workbook.GetWorkSheet("Sheet1");
            //You can also use workbook.DefaultWorkSheet to get default in case you want to get first sheet only
            WorkSheet sheet = workbook.DefaultWorkSheet;
            //Convert the worksheet to System.Data.DataTable
            //Boolean parameter sets the first row as column names of your table.
            return sheet.ToDataTable(true);
        }

        public static void ReadCSVData(string csvFileName)
        {
            var csvFilereader = new DataTable();
            csvFilereader = ReadExcel(csvFileName);


            Console.WriteLine("Enter column no. in which to search, should be a number:");
            int colNo = Convert.ToInt16(Console.ReadLine());

            Console.WriteLine("Enter Key Search Value:");
            string searchValue = Console.ReadLine();


            int totColCount = csvFilereader.Columns.Count;


            if ((totColCount - 1) < colNo)
            {
                Console.WriteLine("\n Error: You entered Invalid Column Number or Column Exceeds the limit");               
            }
            else
            {
                string colName = csvFilereader.Columns[colNo].ColumnName;
                string qry = colName + "='" + searchValue + "\'";


                Console.WriteLine("\n\n Your Search Result is:\n");

                DataRow[] rslt = csvFilereader.Select(qry);
                foreach (DataRow row in rslt)
                {
                    for (int i = 0; i < row.ItemArray.Count(); i++)
                    {
                        Console.WriteLine(i + " | " + row.ItemArray[i]);
                    }
                }
                if (rslt.Length == 0)
                {
                    Console.WriteLine("No Data Found!");
                }
            }

        }
    }
}
