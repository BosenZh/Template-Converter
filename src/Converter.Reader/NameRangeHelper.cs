using SpreadsheetLight;

namespace Converter.Reader
{
    internal class NameRangeHelper
    {
        private readonly SLDocument doc;

        public NameRangeHelper(SLDocument document)
        {
            this.doc = document;
        }

        public string GetWorksheetOfNamedRange(string namedRangeReference)
        {
            if (!string.IsNullOrEmpty(namedRangeReference))
            {
                string[] parts = namedRangeReference.Split('!');
                if (parts.Length > 1)
                {
                    return parts[0].Trim('\'');
                }
            }
            return string.Empty;
        }

        public TableRanges GetTableRanges(string namedRange)
        {
            List<SLDefinedName> definedNames = this.doc.GetDefinedNames();

            SLDefinedName? definedName = definedNames.Find(dn => dn.Name.Equals(namedRange, StringComparison.OrdinalIgnoreCase))
                ?? throw new ArgumentException("The specified named range does not exist.");

            string rangeReference = definedName.Text;

            string[] parts = rangeReference.Split('!');
            if (parts.Length != 2)
            {
                throw new ArgumentException("Invalid named range format.");
            }

            string cellRange = parts[1];
            string startCell = string.Empty;
            string endCell = string.Empty;

            if (cellRange.Contains(':'))
            {
                startCell = cellRange.Split(':')[0];
                endCell = cellRange.Split(':')[1];
                endCell = endCell.Replace("$", string.Empty);
            }
            else
            {
                startCell = cellRange;
            }

            startCell = startCell.Replace("$", string.Empty);

            // Convert the start cell to row and column indices
            (int startRow, int startColumn) = SplitCellReference(startCell);
            (int endRow, int endColumn) = !string.IsNullOrEmpty(endCell)
                ? SplitCellReference(endCell)
                : (startRow, startColumn);
            return new TableRanges()
            {
                StartRow = startRow,
                StartColumn = startColumn,
                EndRow = endRow,
                EndColumn = endColumn,
            };
        }

        private static (int Row, int Column) SplitCellReference(string cellReference)
        {
            string columnLetters = string.Empty;
            string rowNumbers = string.Empty;

            foreach (char c in cellReference)
            {
                if (char.IsLetter(c))
                {
                    columnLetters += c;
                }
                else if (char.IsDigit(c))
                {
                    rowNumbers += c;
                }
            }

            int column = ColumnLetterToNumber(columnLetters);
            int row = int.Parse(rowNumbers);

            return (row, column);
        }

        private static int ColumnLetterToNumber(string columnLetters)
        {
            int sum = 0;
            foreach (char c in columnLetters)
            {
                sum *= 26;
                sum += c - 'A' + 1;
            }

            return sum;
        }
    }
}
