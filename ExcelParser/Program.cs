using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParser
{
    class Program
    {
        static void Main(string[] args)
        {
            var variable = args[0];
            //Console.WriteLine("Type path of excel file:\n");
            string path = @variable;

            Log.Instance.LogPath = @"C:\Users\Bianca\Desktop\monitorLog";
            Log.Instance.LogFileName = "LogFile";

            Log.WriteLine(DateTime.Now + " " + "Creating new data reader.");
            var dataReader = new ExcelDataReader(path);
            Log.WriteLine(DateTime.Now + " " + "New data reader created.");

            Log.WriteLine(DateTime.Now + " " + "Starting excel data reader.");
            var dictionary = dataReader.getDataEntry();

            Log.WriteLine(DateTime.Now + " " + "Spreadsheet data has been added to dictionary successfully.");

            foreach (KeyValuePair<string, DToolsRow> pair in dictionary)
            {

                /*
                Console.WriteLine(pair.Key.ToString() + " " +
                    //pair.Value.First + " "
                    pair.Value.ProductSKU + " "
                    + pair.Value.ProductManufacturer + " "
                    + pair.Value.ProductModel + " "
                    + pair.Value.ProductDescription + " "
                    + pair.Value.ProductClientDescription + " "
                    + pair.Value.ProductQuantity + " "
                    + pair.Value.ProductCost + " "
                    + pair.Value.ProductMargin + " "
                    + pair.Value.ProductMarkup + " "
                    + pair.Value.ProductPrice + " "
                    + pair.Value.NetUnitPrice + " "
                    + pair.Value.Discount + " "
                    + pair.Value.ProductLocation + " "
                    + pair.Value.ProductSubSystem + " "
                    + pair.Value.ProductWireLength + " "
                    + pair.Value.ProductInstallationPhase + " "
                    + pair.Value.ProductRackMount + " "
                    + pair.Value.Type);
                    */
            }
            Console.WriteLine("succeeded");
            Console.ReadLine();
        }

    }
    
} 