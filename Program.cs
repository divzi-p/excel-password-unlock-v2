using excelapi.logic;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddSingleton<ExcelUnlock>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

var summaries = new[]
{
    "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
};

app.MapGet("/readExcelXLSB", async (ExcelUnlock excelUnlock) =>
{
    
    string storageConnectionString="<connecting string>";
    string blobContainerName="test1";
    string blobName="sampledata1.xlsb";
    string excelPassword="password";
    var result= await excelUnlock.ReadXLSBFileFromBlob(storageConnectionString,blobContainerName,blobName, excelPassword);
    return result;
})
.WithName("ReadProtectedExcelBinaryBlob")
.WithOpenApi();

app.MapGet("/readExcelXLSX", async (ExcelUnlock excelUnlock) =>
{
    
    string storageConnectionString="<connecting string>";
    string blobContainerName="test1";
    string blobName="SampleData-xlsx.xlsx";
    string excelPassword="password";
    //working for reading xlsx file from blob.Change blob name above
    var result= await excelUnlock.ReadProtectedExcelFileFromBlob(storageConnectionString,blobContainerName, blobName, excelPassword);
    return result;
})
.WithName("ReadProtectedExcelBlob")
.WithOpenApi();



app.Run();

