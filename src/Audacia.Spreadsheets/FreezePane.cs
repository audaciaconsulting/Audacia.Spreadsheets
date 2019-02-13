using DocumentFormat.OpenXml.Spreadsheet;

namespace Audacia.Spreadsheets
{
    public class FreezePane
    {
        public string StartingCell { get; set; } = "A2";
        public double FrozenColumns { get; set; } = 0D;
        public double FrozenRows { get; set; } = 1D;

        public void Write(SheetView sheetView)
        {
            var hasFrozenColumns = FrozenColumns > 0D;
            var hasFrozenRows = FrozenRows > 0D;
            
            if (hasFrozenRows || hasFrozenColumns)
            {
                // Assume frozen rows
                var activePane = PaneValues.BottomLeft;
                
                if (hasFrozenRows && hasFrozenColumns)
                {
                    activePane = PaneValues.BottomRight;
                }
                else if (hasFrozenColumns)
                {
                    activePane = PaneValues.TopRight;
                }

                var pane = new Pane
                { 
                    HorizontalSplit = FrozenColumns,
                    VerticalSplit = FrozenRows,
                    TopLeftCell = StartingCell,
                    ActivePane = activePane,
                    State = PaneStateValues.Frozen
                };
                
                var selection = new Selection { Pane = activePane };
                
                sheetView.Append(pane);
                sheetView.Append(selection);
            }
        }
    }
}