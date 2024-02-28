using Microsoft.SharePoint.Client;
using System.Security;
using OfficeOpenXml;
using System.Text.Json;
using System.Security.Claims;
using System.Security.Principal;
using Azure.Identity;
using System.IO;
using Azure.Storage.Blobs;
using System.Data;
using ExcelDataReader;
using Aspose.Cells;
using System.Data.SqlClient;
//Microsoft.Sharepoint.Client.Online.CSOM
namespace excelapi.logic
{
public class ExcelUnlock
{

public async Task<string> ReadXLSBFileFromBlob(string storageConnectionString, string blobContainerName,string blobName, string excelPassword)
{
    // Create a BlobServiceClient object
    BlobServiceClient blobServiceClient = new BlobServiceClient(storageConnectionString);

    // Get the blob container client
    BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(blobContainerName);

    // Get the blob client
    BlobClient blobClient = containerClient.GetBlobClient(blobName);

    // Download the blob to a MemoryStream
    MemoryStream memoryStream = new MemoryStream();


    await blobClient.DownloadToAsync(memoryStream);

    /// Load the XLSB file
    LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsb);
    loadOptions.Password = excelPassword;
    Workbook workbook = new Workbook(memoryStream, loadOptions);
    Worksheet worksheet = workbook.Worksheets[0];

    bool hasHeader = true; // adjust it accordingly
    List<Dictionary<string, string>> rows = new List<Dictionary<string, string>>();
    var headerRow = worksheet.Cells.GetRow(0);
    var startRow = hasHeader ? 1 : 0;
    for (var rowNum = startRow; rowNum <= worksheet.Cells.MaxDataRow; rowNum++)
    {
        var wsRow = worksheet.Cells.GetRow(rowNum);
        Dictionary<string, string> row = new Dictionary<string, string>();
        for (int colIndex = 0; colIndex <= worksheet.Cells.MaxDataColumn; colIndex++)
        {
            string header = headerRow.GetCellOrNull(colIndex)?.StringValue ?? $"Column {colIndex + 1}";
            string cellValue = wsRow.GetCellOrNull(colIndex)?.StringValue ?? string.Empty;
            row[header] = cellValue;
        }
        rows.Add(row);
    }

    // Convert the list of dictionaries to JSON
    string json = JsonSerializer.Serialize(rows);

    return json;
}

public async Task<string> ReadXLSBFiletoSQL(string storageConnectionString, string blobName, string excelPassword)
{
    // Create a BlobServiceClient object
    BlobServiceClient blobServiceClient = new BlobServiceClient(storageConnectionString);

    // Get the blob container client
    BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient("test1");

    // Get the blob client
    BlobClient blobClient = containerClient.GetBlobClient(blobName);

    // Download the blob to a MemoryStream
    MemoryStream memoryStream = new MemoryStream();
    Console.WriteLine("Memory stream is read" );

  
    await blobClient.DownloadToAsync(memoryStream);

/// Load the XLSB file
   LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsb);
    loadOptions.Password = excelPassword;
    Workbook workbook = new Workbook(memoryStream, loadOptions);
    Worksheet worksheet = workbook.Worksheets[0];

    bool hasHeader = true; // adjust it accordingly
    List<Dictionary<string, string>> rows = new List<Dictionary<string, string>>();
    var headerRow = worksheet.Cells.GetRow(0);
   // var startRow = hasHeader ? 1 : 0;

    DataTable dt = new DataTable();
    
    // Assuming the first row of the worksheet contains the column names
    for (int colIndex = 0; colIndex <= worksheet.Cells.MaxDataColumn; colIndex++)
    {
        string header = worksheet.Cells[0, colIndex].StringValue;
        dt.Columns.Add(header);
    }
int pageSize = 5000;

for (int pageNum = 0; pageNum * pageSize <= worksheet.Cells.MaxDataRow; pageNum++)
{
    int startRow = pageNum * pageSize + 1;
    int endRow = Math.Min((pageNum + 1) * pageSize, worksheet.Cells.MaxDataRow);
    // Add rows to the DataTable
    for (var rowNum = startRow; rowNum <= endRow; rowNum++)
    {
        var row = worksheet.Cells.GetRow(rowNum);
        DataRow dataRow = dt.NewRow();
        for (int colIndex = 0; colIndex <= worksheet.Cells.MaxDataColumn; colIndex++)
        {
            dataRow[colIndex] = row.GetCellOrNull(colIndex)?.StringValue ?? string.Empty;
        }
        dt.Rows.Add(dataRow);
    }

    // Insert the data into SQL Azure.

    using (SqlConnection conn = new SqlConnection("<sqlconnection"))
    {
        conn.Open();
        using (SqlBulkCopy sbc = new SqlBulkCopy(conn))
        {
            sbc.DestinationTableName = "SalesBonus";
            sbc.WriteToServer(dt);
        }
    }
    dt.Clear();
}  
    return "Ok";
 
}

public async Task<string> ReadProtectedExcelFileFromBlob(string storageConnectionString,string blobContainerName, string blobName, string excelPassword)
{
    // Create a BlobServiceClient object
    BlobServiceClient blobServiceClient = new BlobServiceClient(storageConnectionString);

    // Get the blob container client
    BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(blobContainerName);

    // Get the blob client
    BlobClient blobClient = containerClient.GetBlobClient(blobName);

    // Download the blob to a MemoryStream
    MemoryStream memoryStream = new MemoryStream();
    await blobClient.DownloadToAsync(memoryStream);

    // Load the Excel file
    using (ExcelPackage package = new ExcelPackage(memoryStream, excelPassword))
    {
        DataTable dt = new DataTable();
        var worksheet = package.Workbook.Worksheets[0];
        bool hasHeader = true; // adjust it accordingly
        List<Dictionary<string, string>> rows = new List<Dictionary<string, string>>();
        var headerRow = worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column];
        var startRow = hasHeader ? 2 : 1;
        for (var rowNum = startRow; rowNum <= worksheet.Dimension.End.Row; rowNum++)
        {
            var wsRow = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
            Dictionary<string, string> row = new Dictionary<string, string>();
            foreach (var cell in wsRow)
            {
               string header = headerRow[1,cell.Start.Column].Text;
               row[header] = cell.Text;
            }
            rows.Add(row);
        }

   
        // Convert the DataTable to JSON
        string json = JsonSerializer.Serialize(rows);

        return json;
    }
}


}

}


