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
            // Generate the column data if someone hasn't filled it in themselves.
            Columns = Extensions.Tables.GetColumns<TEntity>().ToList();
            Data = source;
            IncludeHeaders = true;
        }

        public IEnumerable<TEntity> Data { get; set; } = new List<TEntity>();

        public virtual TableRow FromEntity(TEntity entity)
        {
            // Generate the row if someone hasn't defined it themselves.
            return entity.GetRow(Columns);
        }

        public override CellReference Write(SharedDataTable sharedData, OpenXmlWriter writer)
        {
            var rowReference = new CellReference(StartingCellRef);

            // Write SubTotal above headers.
            WriteSubTotalHeaders(sharedData, writer, rowReference);

            // Write Headers above Data.
            WriteHeaders(sharedData, writer, rowReference);
            WriteDataRows(sharedData, writer, rowReference);

            // Return the cell ref at end of the table.
            return rowReference;
        }

        /// <summary>
        /// Enumerate over all rows and write them using an OpenXmlWriter
        /// This puts them into a MemoryStream, to improve this we would
        /// need to update the OpenXml library which we use.
        /// </summary>
        private void WriteDataRows(SharedDataTable sharedData, OpenXmlWriter writer, CellReference rowReference)
        {
            foreach (var entity in Data)
            {
                var row = FromEntity(entity);
                var rowCellReference = rowReference.Clone();
                row.Write(rowCellReference, Columns, sharedData, writer);
                rowReference.NextRow();
            }
        }

        private void WriteHeaders(SharedDataTable sharedData, OpenXmlWriter writer, CellReference rowReference)
        {
            if (IncludeHeaders)
            {
                var headerCellReference = rowReference.Clone();
                var row = new Row();
                writer.WriteStartElement(row);

                foreach (var column in Columns)
                {
                    var isFirstColumn = column == Columns.ElementAt(0);
                    var isLastColumn = column == Columns.ElementAt(Columns.Count - 1);
                    column.Write(HeaderStyle, headerCellReference, isFirstColumn, isLastColumn, sharedData, writer);
                    headerCellReference.NextColumn();
                }

                writer.WriteEndElement();
                rowReference.NextRow();
            }
        }

        private void WriteSubTotalHeaders(SharedDataTable sharedData, OpenXmlWriter writer, CellReference rowReference)
        {
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