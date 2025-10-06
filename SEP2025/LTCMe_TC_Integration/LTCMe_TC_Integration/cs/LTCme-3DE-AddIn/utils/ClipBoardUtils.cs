using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace DemoAddInTC.utils
{
    public class ClipboardUtils
    {
        public static void OnDataGridViewPaste(object grid, KeyEventArgs e)
        {
            if ((e.Shift && e.KeyCode == Keys.Insert) || (e.Control && e.KeyCode == Keys.V))
            {
                PasteTSV((DataGridView)grid);
            }
        }

        public static void PasteTSV(DataGridView grid)
        {
            char[] rowSplitter = { '\r', '\n' };
            char[] columnSplitter = { '\t' };

            // Get the text from clipboard
            IDataObject dataInClipboard = Clipboard.GetDataObject();
            string stringInClipboard = (string)dataInClipboard.GetData(DataFormats.Text);

            // Split it into lines
            string[] rowsInClipboard = stringInClipboard.Split(rowSplitter, StringSplitOptions.RemoveEmptyEntries);

            // Get the row and column of selected cell in grid
            int r = grid.SelectedCells[0].RowIndex;
            int c = grid.SelectedCells[0].ColumnIndex;

            // Add rows into grid to fit clipboard lines
            if (grid.Rows.Count < (r + rowsInClipboard.Length))
            {
                grid.Rows.Add(r + rowsInClipboard.Length - grid.Rows.Count);
            }

            // Loop through the lines, split them into cells and place the values in the corresponding cell.
            for (int iRow = 0; iRow < rowsInClipboard.Length; iRow++)
            {
                // Split row into cell values
                string[] valuesInRow = rowsInClipboard[iRow].Split(columnSplitter);

                // Cycle through cell values
                for (int iCol = 0; iCol < valuesInRow.Length; iCol++)
                {

                    // Assign cell value, only if it within columns of the grid
                    if (grid.ColumnCount - 1 >= c + iCol)
                    {
                        DataGridViewCell cell = grid.Rows[r + iRow].Cells[c + iCol];

                        if (!cell.ReadOnly)
                        {
                            cell.Value = valuesInRow[iCol];
                        }
                    }
                }
            }
        }
    }
}
