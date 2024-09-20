// See https://aka.ms/new-console-template for more information
using ConsoleApp1MAX;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using QuestPDF.Previewer;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Diagnostics;




class Program
{
    static void Main(string[] args)
    {
        var dbManager = new DatabaseManager("localhost", "BL", "SYSADM", "SYSADM");

        List<Items> itemsList = new List<Items>(); // 將 items 改名為 itemsList

        using (var connection = dbManager.OpenConnection())
        {
            string query = "SELECT * FROM Items"; // 從 Items 表中選取所有資料
            using (var command = new SqlCommand(query, connection))
            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    itemsList.Add(new Items // 將 Item 改名為 Items
                    {
                        Id = reader.GetInt32(0),
                        Name = reader.GetString(1),
                        Description = reader.GetString(2),
                        MarketValue = reader.GetInt32(3),
                        Quantity = reader.GetInt32(4),
                        Type = reader.GetString(5),
                        LastUpdated = reader.GetDateTime(6)
                    });
                }
            }
        }

        Document.Create(container =>
        {
            container.Page(Page =>
            {
                // 表頭
                Page.Header()
                .Background(Colors.Green.Lighten2)
                .Text("這是表頭")
                .FontFamily("Microsoft JhengHei");

                // 內文
                Page.Content()
                .DefaultTextStyle(TextStyle.Default.FontFamily("Microsoft JhengHei"))
                .Background(Colors.Blue.Lighten2)
                .AlignCenter()
                .Table(table =>
                {
                    table.ColumnsDefinition(columns =>
                    {
                        columns.ConstantColumn(50); // ID 欄位
                        columns.RelativeColumn(1);   // Name 欄位
                        columns.RelativeColumn(2);   // Description 欄位
                        columns.RelativeColumn(1);   // Market Value 欄位
                        columns.RelativeColumn(1);   // Quantity 欄位
                        columns.RelativeColumn(1);   // Type 欄位
                        columns.RelativeColumn(1);   // Last Updated 欄位
                    });

                    // 表頭
                    table.Header(headers =>
                    {
                        headers.Cell().BorderBottom(1).Text("ID");
                        headers.Cell().BorderBottom(1).Text("Name");
                        headers.Cell().BorderBottom(1).Text("Description");
                        headers.Cell().BorderBottom(1).Text("Market Value");
                        headers.Cell().BorderBottom(1).Text("Quantity");
                        headers.Cell().BorderBottom(1).Text("Type");
                        headers.Cell().BorderBottom(1).Text("Last Updated");
                    });

                    // 填充資料
                    foreach (var item in itemsList) // 將 items 改名為 itemsList
                    {
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten2).Text(item.Id.ToString());
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten2).Text(item.Name);
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten2).Text(item.Description);
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten2).Text(item.MarketValue.ToString());
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten2).Text(item.Quantity.ToString());
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten2).Text(item.Type);
                        table.Cell().BorderBottom(1).BorderColor(Colors.Grey.Lighten2).Text(item.LastUpdated.ToString("yyyy-MM-dd HH:mm:ss"));
                    }
                });

                // 表尾
                Page.Footer()
                .Background(Colors.Purple.Lighten2)
                .Text("這是表尾")
                .FontFamily("Microsoft JhengHei");
            });
        })
        .GeneratePdf("out.pdf");

        // 直接開啟檔案
        Process.Start(new ProcessStartInfo()
        {
            FileName = "out.pdf",
            UseShellExecute = true
        });
    }
}