using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;

namespace ExcelParser
{
    public class OWorkSheet
    {
        private readonly object[,] _workSheetData;
        
        public string Name { get; }

        public OWorkSheet(ExcelWorksheet excelWorksheet)
        {
            _workSheetData = excelWorksheet.Cells.GetValue<object[,]>();
            Name = excelWorksheet.Name;
        }

        public TRow[] ReadAs<TRow>(DataFlow flow = DataFlow.FirstRowAsHeader)
            where TRow : new()
        {
            PropertyInfo[] headersProperties = GetHeaders(flow).Select(GetExcelProperty<TRow>).ToArray();

            return flow == DataFlow.FirstRowAsHeader
                ? ReadAsWithRowAsHeader<TRow>(headersProperties)
                : ReadAsWithCollAsHeader<TRow>(headersProperties);
        }

        #region private methods

        private TRow[] ReadAsWithRowAsHeader<TRow>(PropertyInfo[] headers)
            where TRow : new()
        {
            TRow[] rows = new TRow[_workSheetData.GetLength(0) - 1];

            for (int rowIndex = 1; rowIndex < _workSheetData.GetLength(0); rowIndex++)
            {
                TRow row = new TRow();
                for (int colIndex = 0; colIndex < _workSheetData.GetLength(1); colIndex++)
                {
                    SetValueByHeader(headers, colIndex, row, rowIndex);
                }
                rows[rowIndex - 1] = row;
            }
            return rows;
        }

        private TRow[] ReadAsWithCollAsHeader<TRow>(PropertyInfo[] headers)
            where TRow : new()
        {
            TRow[] rows = new TRow[_workSheetData.GetLength(1) - 1];

            for (int colIndex = 1; colIndex < _workSheetData.GetLength(1); colIndex++)
            {
                TRow row = new TRow();

                for (int rowIndex = 0; rowIndex < _workSheetData.GetLength(0); rowIndex++)
                {
                    SetValueByHeader(headers, colIndex, row, rowIndex);
                }
                rows[colIndex - 1] = row;
            }
            return rows;
        }

        private void SetValueByHeader<TRow>(PropertyInfo[] headers, int colIndex, TRow row, int rowIndex) where TRow : new()
        {
            if (headers[colIndex] != null) // havent got the Excel property
                try
                {
                    headers[colIndex].SetValue(row, _workSheetData[rowIndex, colIndex]);
                }
                catch (InvalidCastException)
                {
                    SetDefaultValue(headers[colIndex], row); //couldnt convert to expected type
                }
        }

        private IEnumerable<string> GetHeaders(DataFlow flow)
        {
            if (flow == DataFlow.FirstRowAsHeader)
                for (int col = 0; col < _workSheetData.GetLength(1); col++)
                {
                    yield return _workSheetData[0, col] as string;
                }
            else
                for (int row = 0; row < _workSheetData.GetLength(0); row++)
                {
                    yield return _workSheetData[row, 0] as string;
                }
        }

        private void SetDefaultValue(PropertyInfo property, object row)
        {
            object defaultValue = null;
            if (property.PropertyType.IsValueType)
                defaultValue = Activator.CreateInstance(property.PropertyType);

            property.SetValue(row, defaultValue);
        }

        private static PropertyInfo GetExcelProperty<TRow>(string header)
        {
            return typeof(TRow)
                .GetProperties(BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance)
                .FirstOrDefault(prop => prop.GetCustomAttribute<ExcelProperty>().ExcelHeader == header);
        }

        #endregion
    }
}