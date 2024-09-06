using System;
using System.Data.SqlClient;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

class Program
{
    static void Main(string[] args)
    {

        string directoryPath = "C:\\Users\\vippr\\source\\repos\\ReadExcelFiles\\ReadExcelFiles\\ReadExcel\\";

        DirectoryInfo di = new DirectoryInfo(directoryPath);

        

        string excelFilePath = @"C:\Path\To\Your\ExcelFile.xlsx";
        string connectionString = "Server=YourServerName;Database=YourDatabaseName;Trusted_Connection=True;";

        try
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelFilePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart!;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    // Skip the header row
                    foreach (Row row in sheetData.Elements<Row>().Skip(1))
                    {
                        string query = "INSERT INTO YourTableName (Column1, Column2, Column3) VALUES (@val1, @val2, @val3)";
                        using (SqlCommand command = new SqlCommand(query, connection))
                        {
                            command.Parameters.AddWithValue("@val1", GetCellValue(workbookPart, row.Elements<Cell>().ElementAtOrDefault(0)));
                            command.Parameters.AddWithValue("@val2", GetCellValue(workbookPart, row.Elements<Cell>().ElementAtOrDefault(1)));
                            command.Parameters.AddWithValue("@val3", GetCellValue(workbookPart, row.Elements<Cell>().ElementAtOrDefault(2)));

                            command.ExecuteNonQuery();
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }

        Console.WriteLine("Data import completed.");
    }

    private static object GetCellValue(WorkbookPart workbookPart, Cell? cell)
    {
        if (cell == null)
            return DBNull.Value;

        string? value = cell.InnerText;

        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            return workbookPart.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>()
                .ElementAt(int.Parse(value)).InnerText;
        }

        return value;
    }
}