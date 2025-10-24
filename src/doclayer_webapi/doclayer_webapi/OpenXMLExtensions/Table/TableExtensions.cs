using D = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;



namespace OpenXMLExtensions
{
    public static class TableExtensions
    {

        /// <summary>
        /// Gets the cell at the specified row and column.
        /// </summary>
        public static D.TableCell GetCell(this D.Table table, int rowNum, int colNum)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));

            var rows = table.Elements<D.TableRow>().ToList();
            if (rowNum <= 0 || rowNum > rows.Count) throw new ArgumentException("Row number not valid");

            var row = rows[rowNum - 1];
            var cols = table.GetFirstChild<D.TableGrid>()?.Elements<D.GridColumn>().ToList();
            if (cols == null || colNum <= 0 || colNum > cols.Count) throw new ArgumentException("Column number not valid");

            return row.Elements<D.TableCell>().ElementAt(colNum - 1);
        }


        /// <summary>
        /// Gets the number of rows in the table.
        /// </summary>
        public static int GetRowCount(this D.Table table)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            return table.Elements<D.TableRow>().Count();
        }

        /// <summary>
        /// Gets the number of columns in the table.
        /// </summary>
        public static int GetColumnCount(this D.Table table)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            var tableGrid = table.GetFirstChild<D.TableGrid>();
            if (tableGrid == null) throw new Exception("No table grid found");

            var columns = tableGrid.Elements<D.GridColumn>().ToList();
            if (columns.Count == 0) throw new Exception("No grid columns found");

            return columns.Count;
        }

        /// <summary>
        /// Adds the specified number of columns to the table.
        /// </summary>
        public static void AddColumns(this D.Table table, int numCols)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (numCols <= 0) throw new ArgumentException("Number of columns must be greater than zero");

            var tableGrid = table.GetFirstChild<D.TableGrid>() ?? new D.TableGrid();
            var colWidth = tableGrid.Elements<D.GridColumn>().LastOrDefault()?.Width ?? 914_400;
            var cols = new List<D.GridColumn>();

            for (int i = 0; i < numCols; i++)
            {
                cols.Add(new D.GridColumn(new D.ExtensionList()) { Width = colWidth });
            }

            foreach (var col in cols)
            {
                tableGrid.AppendChild(col);
            }

            if (table.GetFirstChild<D.TableGrid>() == null)
            {
                table.AppendChild(tableGrid);
            }
        }


        /// <summary>
        /// Adds the specified number of rows to the table.
        /// </summary>
        public static void AddRows(this D.Table table, int numRows)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (numRows <= 0) throw new ArgumentException("Number of rows must be greater than zero");

            var rows = table.Elements<D.TableRow>().ToList();
            if (rows.Count == 0) throw new Exception("Table not valid");

            var numCols = rows.Last().Elements<D.TableCell>().Count();
            var rowHeight = rows.Last().Height;
            var newRows = new List<D.TableRow>();

            for (int i = 0; i < numRows; i++)
            {
                var row = new D.TableRow() { Height = rowHeight };
                for (int k = 0; k < numCols; k++)
                {
                    var tableCell = new D.TableCell();
                    var paragraph = new D.Paragraph();
                    paragraph.AppendChild(new D.EndParagraphRunProperties() { Language = "en-US" });
                    var textBody = new D.TextBody();
                    textBody.AppendChild(new D.BodyProperties());
                    textBody.AppendChild(new D.ListStyle());
                    textBody.AppendChild(paragraph);
                    tableCell.AppendChild(textBody);
                    tableCell.AppendChild(new D.TableCellProperties());
                    row.AppendChild(tableCell);
                }
                newRows.Add(row);
            }

            foreach (var row in newRows)
            {
                table.AppendChild(row);
            }
        }

        /// <summary>
        /// Sets the height of the specified row.
        /// </summary>
        public static void SetRowHeight(this D.Table table, int rowNum, int heightInInches)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (heightInInches <= 0) throw new ArgumentException("Height must be greater than zero");

            var rows = table.Elements<D.TableRow>().ToList();
            if (rowNum <= 0 || rowNum > rows.Count) throw new ArgumentException("Row number not valid");

            rows[rowNum - 1].Height = heightInInches * 914_400;
        }

        /// <summary>
        /// Gets the height of the specified row.
        /// </summary>
        public static Int64Value GetRowHeight(this D.Table table, int rowNum)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));

            var rows = table.Elements<D.TableRow>().ToList();
            if (rowNum <= 0 || rowNum > rows.Count) throw new ArgumentException("Row number not valid");

            return rows[rowNum - 1].Height;
        }


        /// <summary>
        /// Sets the width of the specified column.
        /// </summary>
        public static void SetColumnWidth(this D.Table table, int colNum, int widthInInches)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (widthInInches <= 0) throw new ArgumentException("Width must be greater than zero");

            var tableGrid = table.GetFirstChild<D.TableGrid>();
            if (tableGrid == null) throw new Exception("Table Grid not found");

            var columns = tableGrid.Elements<D.GridColumn>().ToList();
            if (colNum <= 0 || colNum > columns.Count) throw new ArgumentException("Column number not valid");

            columns[colNum - 1].Width = widthInInches * 914_400;
        }

        /// <summary>
        /// Gets the width of the specified column.
        /// </summary>
        public static Int64Value GetColumnWidth(this D.Table table, int colNum)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));

            var tableGrid = table.GetFirstChild<D.TableGrid>();
            if (tableGrid == null) throw new Exception("Table Grid not found");

            var columns = tableGrid.Elements<D.GridColumn>().ToList();
            if (colNum <= 0 || colNum > columns.Count) throw new ArgumentException("Column number not valid");

            return columns[colNum - 1].Width;
        }

        /// <summary>
        /// Sets the table to have no fill.
        /// </summary>
        public static void SetNoFill(this D.Table table)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));

            var props = table.GetFirstChild<D.TableProperties>();
            if (props != null)
            {
                props.AppendChild(new D.NoFill());
            }
        }


        public static void SetRowSchemeFill(this D.Table table, int numRow, int accentNum)
        {
            if (table.Descendants<D.TableCell>().Count() > 0)
            {
                if (numRow <= table.Elements<D.TableRow>().Count() && numRow > 0)
                {
                    D.TableRow row = table.Elements<D.TableRow>().ElementAt(numRow - 1);
                    foreach (D.TableCell cell in row.Elements<D.TableCell>())
                    {
                        cell.SetSchemeFill(accentNum);
                    }
                }
            }
            else
            {
                throw new Exception("Table row not found");
            }
        }

        public static void SetColumnSchemeFill(this D.Table table, int numCol, int accentNum)
        {
            if (table.Descendants<D.TableCell>().Count() > 0)
            {
                if (numCol <= table.Descendants<D.GridColumn>().Count() && numCol > 0)
                {
                    foreach (D.TableRow row in table.Elements<D.TableRow>())
                    {
                        D.TableCell cell = row.Elements<D.TableCell>().ElementAt(numCol - 1);
                        cell.SetSchemeFill(accentNum);
                    }
                }
            }
            else
            {
                throw new Exception("Table column not found");
            }
        }


        public static void SetRowFontColor(this D.Table table, int numRow, int accentNum)
        {
            if (table.Descendants<D.TableCell>().Count() > 0)
            {
                if (numRow <= table.Elements<D.TableRow>().Count() && numRow > 0)
                {
                    D.TableRow row = table.Elements<D.TableRow>().ElementAt(numRow - 1);
                    foreach (D.TableCell cell in row.Elements<D.TableCell>())
                    {
                        cell.SetFontColorScheme(accentNum);
                    }
                }
            }
            else
            {
                throw new Exception("Table row not found");
            }
        }

        public static void SetColumnFontColor(this D.Table table, int numCol, int accentNum)
        {
            if (table.Descendants<D.TableCell>().Count() > 0)
            {
                if (numCol <= table.Descendants<D.GridColumn>().Count() && numCol > 0)
                {
                    foreach (D.TableRow row in table.Elements<D.TableRow>())
                    {
                        D.TableCell cell = row.Elements<D.TableCell>().ElementAt(numCol - 1);
                        cell.SetFontColorScheme(accentNum);
                    }
                }
            }
            else
            {
                throw new Exception("Table column not found");
            }
        }

        public static void SetBottomBorder(this D.Table table, int width)
        {
            if (table.Descendants<D.TableCell>().Count() > 0)
            {
                D.TableRow row = table.Elements<D.TableRow>().Last();
                foreach (D.TableCell cell in row.Elements<D.TableCell>())
                {
                    cell.SetBottomBorder(width);
                }
            }
            else
            {
                throw new Exception("Table row not found");
            }
        }

        public static void SetRowBottomBorder(this D.Table table, int numRow, int width)
        {
            if (table.Descendants<D.TableCell>().Count() > 0)
            {
                if (numRow <= table.Elements<D.TableRow>().Count() && numRow > 0)
                {
                    D.TableRow row = table.Elements<D.TableRow>().ElementAt(numRow - 1);
                    foreach (D.TableCell cell in row.Elements<D.TableCell>())
                    {
                        cell.SetBottomBorder(width);
                    }
                }
                throw new Exception("Row number not valid");

            }
            else
            {
                throw new Exception("Table row not found");
            }
        }

        public static void SetRowTopBorder(this D.Table table, int numRow, int width)
        {
            if (table.Descendants<D.TableCell>().Count() > 0)
            {
                if (numRow <= table.Elements<D.TableRow>().Count() && numRow > 0)
                {
                    D.TableRow row = table.Elements<D.TableRow>().ElementAt(numRow - 1);
                    foreach (D.TableCell cell in row.Elements<D.TableCell>())
                    {
                        cell.SetTopBorder(width);
                    }
                }                 
            }
            else
            {
                throw new Exception("Table row not found");
            }
        }

        public static void SetColumnLeftBorder(this D.Table table, int numCol, int width)
        {
            if (table.Descendants<D.TableCell>().Count() > 0)
            {
                if (numCol <= table.Descendants<D.GridColumn>().Count() && numCol > 0)
                {
                    foreach (D.TableRow row in table.Elements<D.TableRow>())
                    {
                        D.TableCell cell = row.Elements<D.TableCell>().ElementAt(numCol - 1);
                        cell.SetLeftBorder(width);
                    }
                }                    
            }
            else
            {
                throw new Exception("Table cell not found");
            }
        }

        public static void SetColumnRightBorder(this D.Table table, int numCol, int width)
        {
            if (table.Descendants<D.TableCell>().Count() > 0)
            {
                if (numCol <= table.Descendants<D.GridColumn>().Count() && numCol > 0)
                {
                    foreach (D.TableRow row in table.Elements<D.TableRow>())
                    {
                        D.TableCell cell = row.Elements<D.TableCell>().ElementAt(numCol - 1);
                        cell.SetRightBorder(width);
                    }
                }
                else throw new Exception("Column number not valid");                    
            }
            else
            {
                throw new Exception("Table cell not found");
            }
        }

        public static void SetRowBottomBorderColorScheme(this D.Table table, int numRow, int accentNum)
        {
            if (table.Descendants<D.TableCell>().Count() > 0)
            {
                if (numRow <= table.Elements<D.TableRow>().Count() && numRow > 0)
                {
                    D.TableRow row = table.Elements<D.TableRow>().ElementAt(numRow - 1);
                    foreach (D.TableCell cell in row.Elements<D.TableCell>())
                    {
                        cell.SetBottomBorderColorScheme(accentNum);
                    }
                }                    
            }
            else
            {
                throw new Exception("Table row not found");
            }
        }

        public static void SetRowTopBorderColorScheme(this D.Table table, int numRow, int accentNum)
        {
            if (table.Descendants<D.TableCell>().Count() > 0)
            {
                if (numRow <= table.Elements<D.TableRow>().Count() && numRow > 0)
                {
                    D.TableRow row = table.Elements<D.TableRow>().ElementAt(numRow - 1);
                    foreach (D.TableCell cell in row.Elements<D.TableCell>())
                    {
                        cell.SetTopBorderColorScheme(accentNum);
                    }
                }
                else new Exception("Row number not valid");
            }
            else
            {
                throw new Exception("Table row not found");
            }
        }

        public static void SetColumnLeftBorderColorScheme(this D.Table table, int numCol, int accentNum)
        {
            if (table.Descendants<D.TableCell>().Count() > 0)
            {
                if (numCol <= table.Descendants<D.GridColumn>().Count() && numCol > 0)
                    {
                    foreach (D.TableRow row in table.Elements<D.TableRow>())
                    {
                        D.TableCell cell = row.Elements<D.TableCell>().ElementAt(numCol - 1);
                        cell.SetLeftBorderColorScheme(accentNum);
                    }
                }
                    
            }
            else
            {
                throw new Exception("Table cell not found");
            }
        }

        public static void SetRightBorderColorScheme(this D.Table table, int numCol, int accentNum)
        {
            if (table.Descendants<D.TableCell>().Count() > 0)
            {
                if (numCol <= table.Descendants<D.GridColumn>().Count() && numCol > 0)
                {
                    foreach (D.TableRow row in table.Elements<D.TableRow>())
                    {
                        D.TableCell cell = row.Elements<D.TableCell>().ElementAt(numCol - 1);
                        cell.SetRightBorderColorScheme(accentNum);
                    }
                }                   
            }
            else
            {
                throw new Exception("Table cell not found");
            }
        }


        public static void SetRowBottomBorderColorHex(this D.Table table, int numRow, string rgbColorHex)
        {
            if (table.Descendants<D.TableCell>().Count() > 0)
            {
                if (numRow <= table.Elements<D.TableRow>().Count() && numRow > 0)
                {
                    D.TableRow row = table.Elements<D.TableRow>().ElementAt(numRow - 1);
                    foreach (D.TableCell cell in row.Elements<D.TableCell>())
                    {
                        cell.SetBottomBorderColorHex(rgbColorHex);
                    }
                }                    
            }
            else
            {
                throw new Exception("Table row not found");
            }
        }

        public static void SetRowTopBorderColorHex(this D.Table table, int numRow, string rgbColorHex)
        {
            if (table.Descendants<D.TableCell>().Count() > 0)
            {
                if (numRow <= table.Elements<D.TableRow>().Count() && numRow > 0)
                {
                    D.TableRow row = table.Elements<D.TableRow>().ElementAt(numRow - 1);
                    foreach (D.TableCell cell in row.Elements<D.TableCell>())
                    {
                        cell.SetTopBorderColorHex(rgbColorHex);
                    }
                   }                    
            }
            else
            {
                throw new Exception("Table row not found");
            }
        }

        public static void SetColumnLeftBorderColorHex(this D.Table table, int numCol, string rgbColorHex)
        {
            if (table.Descendants<D.TableCell>().Count() > 0)
            {
                foreach (D.TableRow row in table.Elements<D.TableRow>())
                {
                    D.TableCell cell = row.Elements<D.TableCell>().ElementAt(numCol - 1);
                    cell.SetLeftBorderColorHex(rgbColorHex);
                }
            }
            else
            {
                throw new Exception("Table cell not found");
            }
        }

        public static void SetRightBorderColorHex(this D.Table table, int numCol, string rgbColorHex)
        {
            if (table.Descendants<D.TableCell>().Count() > 0)
            {
                if (numCol <= table.Descendants<D.GridColumn>().Count() && numCol > 0)
                {
                    foreach (D.TableRow row in table.Elements<D.TableRow>())
                    {
                        D.TableCell cell = row.Elements<D.TableCell>().ElementAt(numCol - 1);
                        cell.SetRightBorderColorHex(rgbColorHex);
                    }
                }
                else throw new Exception("Column number not valid");
            }
            else
            {
                throw new Exception("Table cell not found");
            }
        }

        public static void SetTextToCell(this D.Table table, string text, int numRow, int numCol)
        {
            if (table.Descendants < D.TableCell>().Count()>0
                && numRow > 0
                && numRow <= table.Elements<D.TableRow>().Count()
                && numCol > 0
                && numCol <= table.Descendants<D.GridColumn>().Count())
            {
                D.TableRow row = table.Elements<D.TableRow>().ElementAt(numRow - 1);
                D.TableCell cell = row.Elements<D.TableCell>().ElementAt(numCol - 1);
                cell.SetText(text);
            }            
        }

        public static void AddTextToCell(this D.Table table, string text, int numRow, int numCol)
        {
            if (table.Descendants < D.TableCell>().Count()>0
                && numRow > 0
                && numRow <= table.Elements<D.TableRow>().Count()
                && numCol > 0
                && numCol <= table.Descendants<D.GridColumn>().Count())
            {
                D.TableRow row = table.Elements<D.TableRow>().ElementAt(numRow - 1);
                D.TableCell cell = row.Elements<D.TableCell>().ElementAt(numCol - 1);
                cell.AddParagraph(text);
            }
        }

        public static void AddCellTextAt(this D.Table table, string text, int numRow, int numCol, int pos)
        {
            if (table.Descendants<D.TableCell>().Count() > 0
                && numRow > 0
                && numRow <= table.Elements<D.TableRow>().Count()
                && numCol > 0
                && numCol <= table.Descendants<D.GridColumn>().Count())
            {
                D.TableRow row = table.Elements<D.TableRow>().ElementAt(numRow - 1);
                D.TableCell cell = row.Elements<D.TableCell>().ElementAt(numCol - 1);
                cell.AddParagraphAt(text, pos);
            }
        }

        public static void DeleteAllTextInCell(this D.Table table, int numRow, int numCol)
        {
            if (table.Descendants<D.TableCell>().Count() > 0
                && numRow > 0
                && numRow <= table.Elements<D.TableRow>().Count()
                && numCol > 0
                && numCol <= table.Descendants<D.GridColumn>().Count())
            {
                D.TableRow row = table.Elements<D.TableRow>().ElementAt(numRow - 1);
                D.TableCell cell = row.Elements<D.TableCell>().ElementAt(numCol - 1);
                cell.DeleteAllText();
            }
        }

        public static bool TryGetText(this D.Table table, string text, out D.Table tableContainingText)
        {
            D.Run run;
            foreach(D.TableCell cell in table.Descendants<D.TableCell>())
            {
                if (cell.TryGetRunContaining(text, out run))
                {
                    // if the text is there in the table, return the table
                    tableContainingText = table;
                    return true;
                }
            }
            tableContainingText = null;
            return false;
        }

        public static bool TryGetRunContainingText(this D.Table table, string text, out D.Run runContainingText)
        {
            D.Run run;
            foreach (D.TableCell cell in table.Descendants<D.TableCell>())
            {
                if (cell.TryGetRunContaining(text, out run))
                {
                    // if the text is there in the table, return the table
                    runContainingText = run;
                    return true;
                }
            }

            runContainingText = null;
            return false;
        }

        /// <summary>
        /// Returns all the inner text of the table concatenated to one string
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public static string GetAllInnerText(this D.Table table)
        {
            string innerText = "";

            foreach (D.TableCell cell in table.Descendants<D.TableCell>())
            {
                innerText += cell.GetText();

            }

            return innerText;
        }

        public static bool TryGetCellContaining(this D.Table table, string text, out D.TableCell cellContainingText)
        {
            D.Run run;
            foreach (D.TableCell cell in table.Descendants<D.TableCell>())
            {
                if (cell.TryGetRunContaining(text, out run))
                {
                    // if the text is there in the table, return the table
                    cellContainingText = cell;
                    return true;
                }
            }

            cellContainingText = null;
            return false;
        }
        public static void BoldSelectedText(this D.Table table, string text)
        {
            D.Run run;
            foreach ( D.TableCell cell in table.Descendants<D.TableCell>())
            {
                if (cell.TryGetRunContaining(text, out run))
                {
                    run.SetRunBold();
                }                
            }
        }

        public static void ItalicizeSelectedText(this D.Table table, string text)
        {
            D.Run run;
            foreach (D.TableCell cell in table.Descendants<D.TableCell>())
            {
                if (cell.TryGetRunContaining(text, out run))
                {
                    run.SetRunItalic();
                }                
            }
        }

        public static void SetSelectedTextSchemeColor(this D.Table table, string text, int accentNum)
        {
            D.Run run;

            foreach (D.TableCell cell in table.Descendants<D.TableCell>())
            {
                if (cell.TryGetRunContaining(text, out run))
                {
                    run.SetRunSchemeFill(accentNum);
                }
            }
        }

    }
}
