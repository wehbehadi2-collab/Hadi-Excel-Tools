
using System;
using System.Collections.Generic;

namespace Hadi.Excel
{
    public class TableData
    {
        public string WorksheetName { get; private set; }

        public string RangeAddress { get; private set; }

        // Column names taken from the first row of the selection.
        public List<string> Headers { get; private set; }

        // Data rows only. The header row is not included.
        public List<List<object>> Rows { get; private set; }

        public int ColumnCount
        {
            get { return Headers.Count; }
        }

        public int RowCount
        {
            get { return Rows.Count; }
        }

        public TableData(
            string worksheetName,
            string rangeAddress,
            List<string> headers,
            List<List<object>> rows)
        {
            if (headers == null)
                throw new ArgumentNullException(nameof(headers));

            if (rows == null)
                throw new ArgumentNullException(nameof(rows));

            WorksheetName = worksheetName ?? "";
            RangeAddress = rangeAddress ?? "";
            Headers = headers;
            Rows = rows;
        }

        // Finds a column by its header, ignoring case and
        // leading/trailing spaces. Returns -1 if not found.
        public int FindColumn(string header)
        {
            if (string.IsNullOrWhiteSpace(header))
                return -1;

            for (int i = 0; i < Headers.Count; i++)
            {
                if (string.Equals(
                    Headers[i].Trim(),
                    header.Trim(),
                    StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        // Data row and column indexes are zero-based.
        public object GetValue(int rowIndex, int columnIndex)
        {
            if (rowIndex < 0 || rowIndex >= Rows.Count)
                throw new ArgumentOutOfRangeException(nameof(rowIndex));

            if (columnIndex < 0 || columnIndex >= Headers.Count)
                throw new ArgumentOutOfRangeException(nameof(columnIndex));

            return Rows[rowIndex][columnIndex];
        }

        public object GetValue(int rowIndex, string header)
        {
            int columnIndex = FindColumn(header);

            if (columnIndex < 0)
            {
                throw new ArgumentException(
                    "Column not found: " + header,
                    nameof(header));
            }

            return GetValue(rowIndex, columnIndex);
        }
    }
}
