using ClosedXML.Excel;

using Microsoft.AspNetCore.Http.Json;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Reflection;
using System.Text.Json;
using iText.IO.Font.Constants;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.Configure<JsonOptions>(options =>
{
    options.SerializerOptions.IncludeFields = true;
});

var app = builder.Build();
var proxyClient = new HttpClient();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.MapPost("/convert_excel", (Input input_data) =>
{
    var helperClass = new MinimalApiFunctions();
    helperClass.GenerateData(input_data.json_input);

    using (var workbook = new XLWorkbook())
    {
        var worksheet = workbook.Worksheets.Add("Sheet1");
        var currentRow = 1;

        for (int headerIndex = 1; headerIndex <= helperClass.headers.Count(); headerIndex++)
        {
            worksheet.Cell(currentRow, headerIndex).Value = helperClass.headers[headerIndex - 1];
        }

        for (int i = 0; i < helperClass.parsedData.Count(); i++)
        {
            currentRow++;
            List<object> rowValues = helperClass.parsedData.ElementAt(i);
            for (int j = 0; j < rowValues.Count(); j++)
            {
                worksheet.Cell(currentRow, j + 1).Value = rowValues[j].ToString();
            }

        }
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        var content = stream.ToArray();
        return Results.File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "output.xlsx");
    }

})
.WithName("ConvertExcel");

app.MapPost("/convert_pdf", (Input input_data) =>
{
    var helperClass = new MinimalApiFunctions();
    helperClass.GenerateData(input_data.json_input);

    using var streamPdf = new MemoryStream();
    PdfDocument pdfDoc = new PdfDocument(new PdfWriter(streamPdf));
    pdfDoc.SetDefaultPageSize(iText.Kernel.Geom.PageSize.A4);
    Document doc = new Document(pdfDoc, iText.Kernel.Geom.PageSize.A4);

    PdfFont fontHeader = PdfFontFactory.CreateFont(StandardFonts.TIMES_BOLD);
    PdfFont fontRows = PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN);
    Table table = new Table(UnitValue.CreatePercentArray(6)).UseAllAvailableWidth();
    table.SetFixedLayout();

    for (int aw = 0; aw < helperClass.headers.Count; aw++)
    {
        Cell cell = new Cell().Add(new Paragraph(helperClass.headers.ElementAt(aw).ToString().ToUpper())
            .SetFont(fontHeader)
            .SetFontColor(ColorConstants.LIGHT_GRAY));
        cell.SetBackgroundColor(ColorConstants.BLUE);
        cell.SetBorder(new SolidBorder(ColorConstants.BLACK, 4));

        cell.SetTextAlignment(TextAlignment.LEFT);
        table.AddCell(cell);
    }
    for (int aw = 0; aw < helperClass.parsedData.Count; aw++)
    {
        List<object> rowValues = helperClass.parsedData.ElementAt(aw);
        for (int j = 0; j < rowValues.Count(); j++)
        {
            Cell cell = new Cell().Add(new Paragraph(rowValues.ElementAt(j).ToString())
                .SetFont(fontRows)
                .SetFontColor(ColorConstants.BLACK));
            cell.SetBackgroundColor(ColorConstants.LIGHT_GRAY);
            cell.SetBorder(Border.NO_BORDER);
            cell.SetBorder(new SolidBorder(ColorConstants.BLACK,2));
            cell.SetTextAlignment(TextAlignment.LEFT);
            table.AddCell(cell);
        }
    }

    doc.Add(table);

    doc.Close();
    return Results.File(streamPdf.ToArray(), "application/pdf", "output.pdf");


})
.WithName("ConvertPdf");

app.Run();

record Input(List<JsonDocument> json_input);

public class MinimalApiFunctions
{
    public List<string> headers = new List<string>();
    public List<List<object>> parsedData = new List<List<object>>();
    //css will be read from a dictionary or smth in future
 

    public void GenerateData(List<JsonDocument> data)
    {
        if (data.Count > 0)
        {
            for (int j = 0; j < data.Count(); j++)
            {
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                dynamic parsedJson = JsonConvert.DeserializeObject(data.ElementAt(j).RootElement.ToString());
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                PropertyInfo[] properties = parsedJson.GetType().GetProperties();
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                for (int i = 0; i < properties.Length; i++)
                {
                    var prop = properties[i];

                    if (prop.Name == nameof(JToken.First) && prop.PropertyType.Name == nameof(JToken))
                    {
                        var token = (JToken)prop.GetValue(parsedJson);
                        List<object> rowData = new List<object>();
                        while (token != null)
                        {
                            if (token is JProperty castProp)
                            {
                                if (j == 0)
                                {
                                    headers.Add(castProp.Name.ToString());
                                }
                                rowData.Add(castProp.Value);
                                Console.WriteLine($"Property: {castProp.Name}; Value: {castProp.Value}");
                            }

                            token = token.Next;
                        }
                        parsedData.Add(rowData);
                    }
                }
            }
        }
    }
}
