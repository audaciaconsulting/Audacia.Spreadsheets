using System.Collections.Generic;
using System.Linq;
using Audacia.Spreadsheets.Extensions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets
{
    public class EntityTable<TEntity> : Table where TEntity : class
    {
        public EntityTable() { }

        public EntityTable(IEnumerable<TEntity> source)
        { 
            // Generate the column data if someone hasn't filled it in themselves...
            Columns = Extensions.Tables.GetColumns<TEntity>();
            Data = source;
            IncludeHeaders = true;
        }

        public IEnumerable<TEntity> Data { get; set; }

        public virtual TableRow FromEntity(TEntity entity)
        {
            // Generate the row if someone hasn't defined it themselves...
            return entity.GetRow(Columns);
        }

        public override CellReference Write(SharedDataTable sharedData, OpenXmlWriter writer)
        {
            var rowReference = new CellReference(StartingCellRef);

            // Write Subtotals above headers
            if (IncludeHeaders && Columns.Any(c => c.DisplaySubtotal))
            {
                var rowCount = Data.Count();
                var subtotalCellRef = rowReference.Clone();
                writer.WriteStartElement(new Row());

                foreach (var column in Columns)
                {
                    var isFirstColumn = column == Columns.ElementAt(0);
                    var isLastColumn = column == Columns.ElementAt(Columns.Count - 1);
                    column.WriteSubtotal(subtotalCellRef, isFirstColumn, isLastColumn, rowCount, sharedData, writer);
                    subtotalCellRef.NextColumn();
                }

                writer.WriteEndElement();
                rowReference.NextRow();
            }

            // Write headers above data
            if (IncludeHeaders)
            {
                var headerCellRef = rowReference.Clone();
                writer.WriteStartElement(new Row());

                foreach (var column in Columns)
                {
                    var isFirstColumn = column == Columns.ElementAt(0);
                    var isLastColumn = column == Columns.ElementAt(Columns.Count - 1);
                    column.Write(HeaderStyle, headerCellRef, isFirstColumn, isLastColumn, sharedData, writer);
                    headerCellRef.NextColumn();
                }

                writer.WriteEndElement();
                rowReference.NextRow();
            }

            // Enumerate over all rows and write them using an openxmlwriter
            // This puts them into a memorystream, to improve this we would need to update the openxml library we are using
            foreach (var entity in Data)
            {
                var row = FromEntity(entity);
                row.Write(rowReference.Clone(), Columns, sharedData, writer);
                rowReference.NextRow();
            }

            // Return the cell ref at end of the table
            return rowReference;
        }
        
        public override int GetMaxCharacterWidth(int columnIndex)
        {
            var column = Columns[columnIndex];

            if (column.Width.HasValue)
            {
                return column.Width.Value;
            }
            
            // TODO: be better another day
            return 75;
        }
    }
}