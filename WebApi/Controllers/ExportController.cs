using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;

namespace WebApi.Controllers;

[ApiController]
[Route("[controller]")]
public class ExportController : ControllerBase
{
    private readonly RandomOrderGenerator randomOrderGenerator;

    public ExportController(RandomOrderGenerator randomOrderGenerator)
    {
        this.randomOrderGenerator = randomOrderGenerator;
    }

    [HttpPost("HelloWorld")]
    public IActionResult Export()
    {
        using var wb = new XLWorkbook();

        var sheet = wb.AddWorksheet("Sheet 1");

        // We can use row and column to select a cell.
        sheet.Cell(1, 1).Value = "Hello";
        // Or we can use cell's address.
        sheet.Cell("B1").Value = "World";

        return SendExcel(wb, "hello-world.xlsx");
    }

    private IActionResult SendExcel(XLWorkbook wb, string filename)
    {
        var stream = new MemoryStream();
        wb.SaveAs(stream);
        stream.Position = 0;

        return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
    }

    [HttpPost("ExportObject")]
    public IActionResult ExportObject()
    {
        using var wb = new XLWorkbook();

        // By default the sheet will be named as 'Sheet 1'.
        var sheet = wb.AddWorksheet();

        // Returns a list of random Order objects of specified count.
        var orders = randomOrderGenerator.Generate(5);

        // Specify a start location at which insertion will take place.
        sheet.Cell(1, 1).InsertData(orders);

        return SendExcel(wb, "orders.xlsx");
    }

    [HttpPost("ExportObjectWithHeader")]
    public IActionResult ExportObjectWithHeader()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet();

        var orders = randomOrderGenerator.Generate(5);

        sheet.Cell(1, 1).InsertTable(orders);

        return SendExcel(wb, "orders-with-header.xlsx");
    }

    [HttpPost("StylingSheet")]
    public IActionResult StylingSheet()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet();

        // Make sure to insert data first and then style
        var orders = randomOrderGenerator.Generate(10);
        sheet.Cell(3, 1).InsertData(orders);

        // Style whole sheet
        sheet.Style.Font.SetFontSize(12);

        // Style rows and columns
        sheet.Columns().Width = 15;
        sheet.Column(4).Width = 25;
        sheet.Rows(1, 2).Style.Fill.BackgroundColor = XLColor.LavenderGray;

        // Style a range
        sheet.Range("A3:D12").Style.Fill.BackgroundColor = XLColor.YellowGreen;

        // Style individual Cell
        sheet.Cell("D12").Style.Font.FontColor = XLColor.Red;

        return SendExcel(wb, "styled-sheet.xlsx");
    }

    [HttpPost("ConditionalFormatting")]
    public IActionResult ConditionalFormatting()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet();

        var orders = randomOrderGenerator.Generate(10);
        sheet.Cell(1, 1).InsertData(orders);

        sheet.Range("C1:C10").AddConditionalFormat()
            .WhenEqualOrGreaterThan(50)
            .Fill.SetBackgroundColor(XLColor.Green);

        sheet.Range("A1:A10").AddConditionalFormat()
            .ColorScale()
            .LowestValue(XLColor.Red)
            .Midpoint(XLCFContentType.Percent, 50, XLColor.Yellow)
            .HighestValue(XLColor.Green);

        return SendExcel(wb, "conditional-formatting.xlsx");
    }

    [HttpPost("AddingHyperlinks")]
    public IActionResult AddingHyperlinks()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet();

        // Set text and hyperlink
        sheet.Cell(2, 1).Value = "Facebook";
        sheet.Cell(2, 1).SetHyperlink(new XLHyperlink("https://www.facebook.com"));

        // Set hyperlink to another cell
        sheet.Cell(3, 1).Value = "Go to B1";
        sheet.Cell(3, 1).SetHyperlink(new XLHyperlink("B1"));

        return SendExcel(wb, "hyperlinks.xlsx");
    }

    [HttpPost("UsingFormulas")]
    public IActionResult UsingFormulas()
    {
        using var wb = new XLWorkbook();
        var sheet = wb.AddWorksheet();

        var orders = randomOrderGenerator.Generate(5);
        sheet.Cell(1, 1).InsertData(orders);

        sheet.Cell("C6").FormulaA1 = "=SUM(C1:C5)";

        return SendExcel(wb, "formulas.xlsx");
    }
}