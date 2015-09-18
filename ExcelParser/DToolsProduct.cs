using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParser
{
    class DToolsProduct
    {
        //public string First { get; set; }
        public string ProductSKU { get; set; }
        public string ProductManufacturer { get; set; }
        public string ProductModel { get; set; }
        public string ProductDescription { get; set; }
        public string ProductClientDescription { get; set; }
        public string ProductQuantity { get; set; }
        public string ProductCost { get; set; }
        public string ProductMargin { get; set; }
        public string ProductMarkup { get; set; }
        public string ProductPrice { get; set; }
        public string NetUnitPrice { get; set; }
        public string Discount { get; set; }
        public string ProductLocation { get; set; }
        public string ProductSubSystem { get; set; }
        public string ProductWireLength { get; set; }
        public string ProductInstallationPhase { get; set; }
        public string ProductRackMount { get; set; }
        public string Type { get; set; }


        public Boolean isEmpty() {
            if (ProductSKU.Length == 0 &&
                ProductManufacturer.Length == 0 &&
                ProductModel.Length == 0 &&
                ProductDescription.Length == 0 &&
                ProductClientDescription.Length == 0 &&
                ProductQuantity.Length == 0 &&
                ProductCost.Length == 0 &&
                ProductMargin.Length == 0 &&
                ProductMarkup.Length == 0 &&
                ProductPrice.Length == 0 &&
                NetUnitPrice.Length == 0 &&
                Discount.Length == 0 &&
                ProductLocation.Length == 0 &&
                ProductSubSystem.Length == 0 &&
                ProductWireLength.Length == 0 &&
                ProductInstallationPhase.Length == 0 &&
                ProductRackMount.Length == 0 &&
                Type.Length == 0
                )
            {
                return true;
            }

            return false;
        }
    }
}
