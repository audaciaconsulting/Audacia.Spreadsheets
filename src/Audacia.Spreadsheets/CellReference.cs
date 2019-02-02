using System;
using Audacia.Spreadsheets.Extensions;

namespace Audacia.Spreadsheets
{
    public class CellReference
    {
        public CellReference()
        {
            RowNumber = 1;
            ColumnLetter = "A";
        }
        
        public CellReference(string cellRef)
        {
            RowNumber = cellRef.GetRowNumber();
            ColumnLetter = cellRef.GetColumnLetter();
        }

        public CellReference(uint rowNumber, string columnLetter)
        {
            RowNumber = rowNumber;
            ColumnLetter = columnLetter;
        }
        
        public static implicit operator string(CellReference source)
        {
            return source.ToString();
        }
        
        /// <summary>
        /// The row number ie; 1, 2, 3, 4.
        /// </summary>
        public uint RowNumber { get; set; }
        
        /// <summary>
        /// The Column Letter ie; A, B, C, D.
        /// </summary>
        public string ColumnLetter { get; set; }

        /// <summary>
        /// Returns a new Object with the same values.
        /// </summary>
        public CellReference Clone()
        {
            return new CellReference(RowNumber, ColumnLetter);
        }
        
        /// <summary>
        /// Returns a new object modified by the given values.
        /// </summary>
        public CellReference MutateBy(int rows, int columns)
        {
            var uintValue = Convert.ToUInt32(rows);
            var nextRowNumber = RowNumber + uintValue;

            if (columns != 0)
            {
                var nextColumnNumber = ColumnLetter.ToColumnNumber() + columns;
                var nextColumnLetter = nextColumnNumber.ToColumnLetter();
                return new CellReference(nextRowNumber, nextColumnLetter);
            }

            return new CellReference(nextRowNumber, ColumnLetter);
        }
        
        /// <summary>
        /// Increments the column value by one.
        /// </summary>
        public void NextColumn()
        {
            var nextColumnNumber = ColumnLetter.ToColumnNumber() + 1;
            ColumnLetter = nextColumnNumber.ToColumnLetter();
        }
        
        /// <summary>
        /// Increments the row value by one.
        /// </summary>
        public void NextRow()
        {
            RowNumber++;
        }
        
        public override string ToString()
        {
            return $"{RowNumber}{ColumnLetter}";
        }
    }
}