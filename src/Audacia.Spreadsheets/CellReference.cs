using System;
using Audacia.Spreadsheets.Extensions;

namespace Audacia.Spreadsheets
{
    public class CellReference
    {
        public CellReference() { }
        
        public CellReference(string cellRef)
        {
            RowIndex = cellRef.GetReferenceRowIndex();
            ColumnIndex = cellRef.GetReferenceColumnIndex();
        }

        public CellReference(uint rowIndex, string columnIndex)
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }
        
        public static implicit operator string(CellReference source)
        {
            return source.ToString();
        }
        
        /// <summary>
        /// The row number ie; 1, 2, 3, 4.
        /// </summary>
        public uint RowIndex { get; set; }
        
        /// <summary>
        /// The Column Letter ie; A, B, C, D.
        /// </summary>
        public string ColumnIndex { get; set; }

        /// <summary>
        /// Returns a new Object with the same values.
        /// </summary>
        public CellReference Clone()
        {
            return new CellReference(RowIndex, ColumnIndex);
        }
        
        /// <summary>
        /// Returns a new object with the rows incremented by the given value.
        /// </summary>
        public CellReference MutateRowsBy(int value)
        {
            var uintValue = Convert.ToUInt32(value);
            return new CellReference(RowIndex + uintValue, ColumnIndex);
        }
        
        /// <summary>
        /// Increments the column value by one.
        /// </summary>
        public void NextColumn()
        {
            var nextColumnNumber = ColumnIndex.GetColumnNumber() + 1;
            ColumnIndex = nextColumnNumber.GetExcelColumnName();
        }
        
        /// <summary>
        /// Increments the row value by one.
        /// </summary>
        public void NextRow()
        {
            RowIndex++;
        }
        
        public override string ToString()
        {
            return $"{RowIndex}{ColumnIndex}";
        }
    }
}