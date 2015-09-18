using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParser
{
    class DToolsRow
    {

        private String rowString;
        private double rowDouble;
        private List<DToolsProduct> rowList;



        public DToolsRow(String _row)
        {
            rowString = _row;

        }

        public DToolsRow(double _row)
        {
            rowDouble = _row;

        }

        public DToolsRow(List<DToolsProduct> _row)
        {
            rowList = _row;

        }


        public string getRowString()
        {
            return rowString;
        }


        public double getRowDouble()
        {
            return rowDouble;
        }

        public List<DToolsProduct> getRowList()
        {
            return rowList;
        }

    }
}
