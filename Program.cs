using DoctorAppointmentWebApi.Messages;
using Microsoft.Extensions.Configuration;
using EasyNetQ;
using iTextSharp.text;
using iTextSharp.text.pdf;

var builder = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
IConfiguration config = builder.Build();
var amqp = config.GetConnectionString("AutoRabbitMQ");
using var bus = RabbitHutch.CreateBus(amqp);
Console.WriteLine("Подключился!");

var subscriberId = $"ReportGenerator@{Environment.MachineName}";

await bus.PubSub.SubscribeAsync<DoctorReportData>(subscriberId, HandleDoctorReportData);

Console.WriteLine("Слушаю RabbitMQ. Жду сообщения");
Console.ReadLine();

void HandleDoctorReportData(DoctorReportData reportData)
{
    Console.WriteLine($"Получил отчётность по доктору: {reportData.DoctorName}");

    CreatePdfReport(reportData);
}

void CreatePdfReport(DoctorReportData reportData)
{
    string fileName = $"DoctorReport_{reportData.DoctorName}_{DateTime.Now:yyyyMMddHHmmss}.pdf";
    string filePath = Path.Combine("Reports", fileName);
    Directory.CreateDirectory("Reports");

    using var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);
    var doc = new Document(PageSize.A4, 20, 20, 30, 30);
    PdfWriter.GetInstance(doc, fs);
    doc.Open();

    var titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 16);
    var title = new Paragraph($"Report for doctor: {reportData.DoctorName}", titleFont);
    title.Alignment = Element.ALIGN_CENTER;
    doc.Add(title);

    doc.Add(new Paragraph($"Specialization: {reportData.Specialization}",
        FontFactory.GetFont(FontFactory.HELVETICA, 12)));
    doc.Add(new Paragraph($"Period: {reportData.Period}", FontFactory.GetFont(FontFactory.HELVETICA, 12)));
    doc.Add(new Paragraph($"All patients: {reportData.TotalPatients}", FontFactory.GetFont(FontFactory.HELVETICA, 12)));
    doc.Add(new Paragraph("\n"));

    var table = new PdfPTable(5) { WidthPercentage = 100 };
    table.SetWidths(new float[] { 2, 2, 2, 2, 4 });

    var headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12);
    var cellFont = FontFactory.GetFont(FontFactory.HELVETICA, 10);

    var headers = new[] { "Name", "Surname", "Date of birth", "Appointment date", "Reports" };
    foreach (var header in headers)
    {
        var cell = new PdfPCell(new Phrase(header, headerFont))
        {
            HorizontalAlignment = Element.ALIGN_CENTER,
            VerticalAlignment = Element.ALIGN_MIDDLE,
            BackgroundColor = new BaseColor(230, 230, 230),
            Padding = 5
        };
        table.AddCell(cell);
    }

    foreach (var patient in reportData.PatientDetails)
    {
        table.AddCell(new PdfPCell(new Phrase(patient.FirstName ?? "Unknown", cellFont))
            { HorizontalAlignment = Element.ALIGN_LEFT, Padding = 5 });
        table.AddCell(new PdfPCell(new Phrase(patient.LastName ?? "Unknown", cellFont))
            { HorizontalAlignment = Element.ALIGN_LEFT, Padding = 5 });
        table.AddCell(new PdfPCell(new Phrase(patient.DateOfBirth?.ToString("dd.MM.yyyy") ?? "N/A", cellFont))
            { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5 });
        table.AddCell(new PdfPCell(new Phrase(patient.AppointmentDate.ToString("dd.MM.yyyy HH:mm"), cellFont))
            { HorizontalAlignment = Element.ALIGN_CENTER, Padding = 5 });
        table.AddCell(new PdfPCell(new Phrase(patient.Symptoms ?? "Unknown", cellFont))
            { HorizontalAlignment = Element.ALIGN_LEFT, Padding = 5 });
    }

    doc.Add(table);
    doc.Close();

    Console.WriteLine($"Отчёт сохранён: {filePath}");
}