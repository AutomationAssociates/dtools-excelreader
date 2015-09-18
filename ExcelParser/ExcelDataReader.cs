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

    class ExcelDataReader
    {
        string _path;

        int counter = 0;


        public ExcelDataReader(string path)
        {
            _path = path;
            Log.WriteLine(DateTime.Now + " " + "File path has been set successfully as: " + _path);
        }

        public IExcelDataReader getExcelReader()
        {
            Log.WriteLine(DateTime.Now + " " + "Attempting to open and read file.");
            FileStream stream = File.Open(_path, FileMode.Open, FileAccess.Read);
            Log.WriteLine(DateTime.Now + " " + "Opened and read file successfully.");


            IExcelDataReader reader = null;
            try
            {
                if (_path.EndsWith(".xlsx"))
                {
                    Log.WriteLine(DateTime.Now + " " + "Preparing reader for .xlsx file.");
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }
                if (_path.EndsWith(".xls"))
                {
                    Log.WriteLine(DateTime.Now + " " + "Preparing reader for .xls file.");
                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                Log.WriteLine(DateTime.Now + " " + "Reader prepared successfully.");
                return reader;
                
            }
            catch (Exception)
            {
                Log.WriteLine(DateTime.Now + " " + "Reader did not succeed to prepare.");
                throw;
            }
        }

        public IEnumerable<string> getWorksheetNames()
        {
            var reader = this.getExcelReader();
            var workbook = reader.AsDataSet();
            Log.WriteLine(DateTime.Now + " " + "Attempting to retrieve sheets from file.");
            var sheets = from DataTable sheet in workbook.Tables select sheet.TableName;
            Log.WriteLine(DateTime.Now + " " + "Retrieved sheets successfully.");
            return sheets;
        }

        public IEnumerable<DataRow> getData(string sheet, bool firstRowIsColumnNames = true)
        {
            var reader = this.getExcelReader();
            reader.IsFirstRowAsColumnNames = firstRowIsColumnNames;
            var workSheet = reader.AsDataSet().Tables[sheet];
            Log.WriteLine(DateTime.Now + " " + "Attempting to retrieve row data from file.");
            var rows = from DataRow a in workSheet.Rows select a;
            Log.WriteLine(DateTime.Now + " " + "Retrieved row data successfully.");
            return rows;
        }

        public Dictionary<string, DToolsRow> getDataEntry()
        {
            Log.WriteLine(DateTime.Now + " " + "Initializing dictionary.");
            Dictionary<string, DToolsRow> dToolsDict = new Dictionary<string, DToolsRow>();

            Log.WriteLine(DateTime.Now + " " + "Attempting to add data to a list.");
            DToolsRow listRow = new DToolsRow(getList());
            Log.WriteLine(DateTime.Now + " " + "Added data to the list successfully.");


            Log.WriteLine(DateTime.Now + " " + "Attempting to add list/first row to dictionary.");
            dToolsDict.Add("Products", listRow);
            Log.WriteLine(DateTime.Now + " " + "Added list as first row to dictionary successfully.");

            Log.WriteLine(DateTime.Now + " " + "Attempting to add string values to the dictionary.");
            dToolsDict = addStringValues(dToolsDict);
            Log.WriteLine(DateTime.Now + " " + "Added string values to the dictionary successfully.");

            return dToolsDict;
        }



        private List<DToolsProduct> getList()
        {

            List<DToolsProduct> dataList = new List<DToolsProduct>();

            var excelData = new ExcelDataReader(_path);
            var data = excelData.getData("Sheet1");


            Log.WriteLine(DateTime.Now + " " + "Attempting to add products to a list.");
            foreach (var row in data)
            {
                

                var dToolsProduct = new DToolsProduct()
                {

                    //First = row["First"].ToString(),
                    ProductSKU = row["Product SKU"].ToString(),
                    ProductManufacturer = row["Product Manufacturer"].ToString(),
                    ProductModel = row["Product Model"].ToString(),
                    ProductDescription = row["Product Description"].ToString(),
                    ProductClientDescription = row["Product Client Description"].ToString(),
                    //Product Quantitiy has to have to spaces as in the file otherwise program doesn't run
                    ProductQuantity = row["Product  Quantity"].ToString(),
                    ProductCost = row["Product Cost"].ToString(),
                    ProductMargin = row["Product Margin"].ToString(),
                    ProductMarkup = row["Product Markup"].ToString(),
                    ProductPrice = row["Product Price"].ToString(),
                    NetUnitPrice = row["Net Unit Price"].ToString(),
                    Discount = row["Discount"].ToString(),
                    ProductLocation = row["Product Location"].ToString(),
                    ProductSubSystem = row["Product Sub-system"].ToString(),
                    //Product Wire Length has to have two spaces as in the file otherwise program doesn't run
                    ProductWireLength = row["Product  Wire Length"].ToString(),
                    //Product Installation Phase has two have two spaces as in the file otherwise program doesn't run
                    ProductInstallationPhase = row["Product Installation Phase  (Prewire etc.)"].ToString(),
                    //Product Rack Mount has to have two spaces as in the file otherwise program doesn't run
                    ProductRackMount = row["Product Rack  Mount Selection"].ToString(),
                    Type = row["Type"].ToString()
                };

                if (dToolsProduct.ProductSKU == "Name") {
                    break;
                }

                if (dToolsProduct.isEmpty() == false) {
                    dataList.Add(dToolsProduct);
                }

                ++counter;
            }
            Log.WriteLine(DateTime.Now + " " + "Added products to the list successfully.");
            return dataList;
        }

        private Dictionary<string, DToolsRow> addStringValues(Dictionary<string, DToolsRow> dToolsDictionary) {

            var excelData = new ExcelDataReader(_path);
            var data = excelData.getData("Sheet1").Skip(counter);
            DToolsRow listRow;


            
            foreach (var row in data)
            {
                
                    
                if (row[1].ToString() == "Name") {
                    
                    var index = row.Table.Rows.IndexOf(row);
                    var array = row.Table.Select();
                    var nextRow = array[index + 1];

                    listRow = new DToolsRow(nextRow["Product SKU"].ToString());
                    dToolsDictionary.Add("Name", listRow);
                    listRow = new DToolsRow(nextRow["Product Manufacturer"].ToString());
                    dToolsDictionary.Add("Description", listRow);
                    listRow = new DToolsRow(nextRow["Product Model"].ToString());
                    dToolsDictionary.Add("Quantity", listRow);
                    listRow = new DToolsRow(nextRow["Product Description"].ToString());
                    dToolsDictionary.Add("UnitPrice", listRow);
                    listRow = new DToolsRow(nextRow["Product Client Description"].ToString());
                    dToolsDictionary.Add("TotalPrice", listRow);


                    Console.WriteLine(index + "next" + nextRow[1].ToString());
                }


                if (row[1].ToString() == "Prewire Hours")
                {

                    var index = row.Table.Rows.IndexOf(row);
                    var array = row.Table.Select();
                    var nextRow = array[index + 1];

                    listRow = new DToolsRow(nextRow["Product SKU"].ToString());
                    dToolsDictionary.Add("prewireHours", listRow);
                    listRow = new DToolsRow(nextRow["Product Manufacturer"].ToString());
                    dToolsDictionary.Add("fitoffHours", listRow);
                    listRow = new DToolsRow(nextRow["Product Model"].ToString());
                    dToolsDictionary.Add("programmingHours", listRow);
                    listRow = new DToolsRow(nextRow["Product Description"].ToString());
                    dToolsDictionary.Add("projectManagementHours", listRow);
                    listRow = new DToolsRow(nextRow["Product Client Description"].ToString());
                    dToolsDictionary.Add("designHours", listRow);
                    listRow = new DToolsRow(nextRow["Product  Quantity"].ToString());
                    dToolsDictionary.Add("miscLabor", listRow);

                    Console.WriteLine(index + "next" + nextRow[1].ToString());
                }


                if (row[1].ToString() == "Project Client Name")
                {

                    var index = row.Table.Rows.IndexOf(row);
                    var array = row.Table.Select();
                    var nextRow = array[index + 1];

                    listRow = new DToolsRow(nextRow[1].ToString());
                    dToolsDictionary.Add("projectClientName", listRow);
                    listRow = new DToolsRow(nextRow[2].ToString());
                    dToolsDictionary.Add("projectName", listRow);
                    listRow = new DToolsRow(nextRow[3].ToString());
                    dToolsDictionary.Add("projectTotal", listRow);
                    listRow = new DToolsRow(nextRow[4].ToString());
                    dToolsDictionary.Add("projectStatus", listRow);
                    listRow = new DToolsRow(nextRow[5].ToString());
                    dToolsDictionary.Add("projectDToolsNumber", listRow);
                    listRow = new DToolsRow(nextRow[6].ToString());
                    dToolsDictionary.Add("projectBrand", listRow);
                    listRow = new DToolsRow(nextRow[7].ToString());
                    dToolsDictionary.Add("projectTotalCost", listRow);
                    listRow = new DToolsRow(Convert.ToDouble(nextRow[8].ToString()));
                    dToolsDictionary.Add("projectTotalSell", listRow);
                    listRow = new DToolsRow(nextRow[9].ToString());
                    dToolsDictionary.Add("projectTotalMiscCosts", listRow);
                    listRow = new DToolsRow(nextRow[10].ToString());
                    dToolsDictionary.Add("projectRevisionNumber", listRow);
                    listRow = new DToolsRow(nextRow[11].ToString());
                    dToolsDictionary.Add("projectProductQty", listRow);
                    listRow = new DToolsRow(nextRow[12].ToString());
                    dToolsDictionary.Add("projectCustProp1", listRow);
                    listRow = new DToolsRow(nextRow[13].ToString());
                    dToolsDictionary.Add("projectCustomProperty2", listRow);
                    listRow = new DToolsRow(nextRow[14].ToString());
                    dToolsDictionary.Add("projectCustomProperty3", listRow);
                    listRow = new DToolsRow(nextRow[15].ToString());
                    dToolsDictionary.Add("projectCustProp4", listRow);
                    listRow = new DToolsRow(nextRow[16].ToString());
                    dToolsDictionary.Add("projectCustomProperty5", listRow);
                    listRow = new DToolsRow(nextRow[17].ToString());
                    dToolsDictionary.Add("projectCustomProperty6", listRow);
                    listRow = new DToolsRow(nextRow[18].ToString());
                    dToolsDictionary.Add("projectCustomProperty7", listRow);
                    listRow = new DToolsRow(nextRow[19].ToString());
                    dToolsDictionary.Add("projectCustomProperty8", listRow);
                    listRow = new DToolsRow(nextRow[20].ToString());
                    dToolsDictionary.Add("projectMiscPartsAdjustment", listRow);
                    listRow = new DToolsRow(nextRow[21].ToString());
                    dToolsDictionary.Add("projectEquipmentAdjustment", listRow);
                    listRow = new DToolsRow(nextRow[22].ToString());
                    dToolsDictionary.Add("projectInstallationHours", listRow);
                    listRow = new DToolsRow(nextRow[23].ToString());
                    dToolsDictionary.Add("projectID", listRow);
                    listRow = new DToolsRow(nextRow[24].ToString());
                    dToolsDictionary.Add("projectDiscount", listRow);
                    listRow = new DToolsRow(nextRow[25].ToString());
                    dToolsDictionary.Add("projectAssigned", listRow);


                    Console.WriteLine(index + "next" + nextRow[1].ToString());
                }

                var indexCurrentRow = row.Table.Rows.IndexOf(row);
                if (indexCurrentRow == (counter + 7)){
                    break;
                }

            }


            return dToolsDictionary;

        }


    }
}
