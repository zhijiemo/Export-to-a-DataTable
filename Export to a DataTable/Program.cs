using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace Export_to_a_DataTable
{
    // Reference:
    // http://www.codeproject.com/Articles/30490/How-to-Manually-Create-a-Typed-DataTable
    

    public class ProductRow : DataRow
    {
        public int ProductID
        {
            get { return (int)base["ProductID"]; }
            set { base["ProductID"] = value; }
        }

        public string ProductDescription
        {
            get { return (string)base["ProductDescription"]; }
            set { base["ProductDescription"] = value; }
        }

        public DateTime DateAdded
        {
            get { return (DateTime)base["DateAdded"]; }
            set { base["DateAdded"] = value; }
        }

        public decimal Price
        {
            get { return (decimal)base["Price"]; }
            set { base["Price"] = value; }
        }

        internal ProductRow(DataRowBuilder builder)
            : base(builder)
        {
            this.ProductID = 0;
            this.ProductDescription = string.Empty;
            this.DateAdded = DateTime.Now;
            this.Price = 0;
        }
    }

    public class ProductTable : DataTable
    {
        public ProductTable(string TableName)
        {
            this.TableName = TableName;
            this.Columns.Add(new DataColumn("ProductID", typeof(int)));
            this.Columns.Add(new DataColumn("ProductDescription", typeof(string)));
            this.Columns.Add(new DataColumn("DateAdded", typeof(DateTime)));
            this.Columns.Add(new DataColumn("Price", typeof(decimal)));
        }


        public ProductRow CreateNewRow()
        {
            return (ProductRow)NewRow();
        }

        protected override Type GetRowType()
        {
            return typeof(ProductRow);
        }

        protected override DataRow NewRowFromBuilder(DataRowBuilder builder)
        {
            return new ProductRow(builder);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            DataTable dtStrongTyping = new DataTable("ProductStrongTyping");
            ProductTable dtSchwarzeneggerTyping = new ProductTable("ProductSchwarzeneggerTyping");

            // this basically makes it equivalent to ProductTable,
            // but we're "manually" doing it in code, instead of having
            // all this done by a "proper" class.
            dtStrongTyping.Columns.Add(new DataColumn("ProductID", typeof(int)));
            dtStrongTyping.Columns.Add(new DataColumn("ProductDescription", typeof(string)));
            dtStrongTyping.Columns.Add(new DataColumn("DateAdded", typeof(DateTime)));
            dtStrongTyping.Columns.Add(new DataColumn("Price", typeof(decimal)));

            int row;
            ProductRow prodrow;

            using (SLDocument sl = new SLDocument("ExportToDataTableExample.xlsx", "Sheet1"))
            {
                SLWorksheetStatistics stats = sl.GetWorksheetStatistics();
                int iStartColumnIndex = stats.StartColumnIndex;

                // I'll assume that the "first" row is always the header row.
                // Notice that the StartRowIndex returned isn't necessarily the
                // first row of the worksheet, it's the first row that has any data.

                // WARNING: the statistics object notes down any non-empty cells.
                // This includes cells with say a background colour but doesn't have
                // cell data. So if you have an empty row just after your worksheet
                // tabular data, but the row is coloured light blue, the EndRowIndex
                // will be one more than you need.
                // It is suggested that you know more about the input Excel file you're
                // using...

                // I'm also not using any variables to store the intermediate returned
                // cell data. This makes each code segment independent of each other,
                // and also makes it such that it's easier for you to see what you
                // actually have to type.

                for (row = stats.StartRowIndex + 1; row <= stats.EndRowIndex; ++row)
                {
                    dtStrongTyping.Rows.Add(
                        sl.GetCellValueAsInt32(row, iStartColumnIndex),
                        sl.GetCellValueAsString(row, iStartColumnIndex + 1),
                        sl.GetCellValueAsDateTime(row, iStartColumnIndex + 2),
                        sl.GetCellValueAsDecimal(row, iStartColumnIndex + 3)
                        );
                }


                for (row = stats.StartRowIndex + 1; row <= stats.EndRowIndex; ++row)
                {
                    prodrow = dtSchwarzeneggerTyping.CreateNewRow();
                    prodrow.ProductID = sl.GetCellValueAsInt32(row, iStartColumnIndex);
                    prodrow.ProductDescription = sl.GetCellValueAsString(row, iStartColumnIndex + 1);
                    prodrow.DateAdded = sl.GetCellValueAsDateTime(row, iStartColumnIndex + 2);
                    prodrow.Price = sl.GetCellValueAsDecimal(row, iStartColumnIndex + 3);
                    dtSchwarzeneggerTyping.Rows.Add(prodrow);
                }
            }
            // just to prove that the data in each DataTable is correct...
            dtStrongTyping.Rows.Add(1024, "I change keyboards every month because they can't handle my strong typing skills", DateTime.Now.AddMonths(1), 2.78m);

            prodrow = dtSchwarzeneggerTyping.CreateNewRow();
            prodrow.ProductID = 2048;
            prodrow.ProductDescription = "I'll be back";
            prodrow.DateAdded = new DateTime(2029, 1, 1);
            prodrow.Price = 78371200m;
            dtSchwarzeneggerTyping.Rows.Add(prodrow);

            // I don't know what you're going to do with the DataTable,
            // so I'm just going to write the contents to XML files.

            dtStrongTyping.WriteXml("ExportToDataTableExampleStrongTyping.xml");
            dtSchwarzeneggerTyping.WriteXml("ExportToDataTableExampleSchwarzeneggerTyping.xml");

            Console.WriteLine("End of program");
            Console.ReadLine();

        }
    }
}
