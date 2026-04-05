using System.IO;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using WBody = DocumentFormat.OpenXml.Wordprocessing.Body;
using WDocument = DocumentFormat.OpenXml.Wordprocessing.Document;
using WParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using WRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using WText = DocumentFormat.OpenXml.Wordprocessing.Text;
using WTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using WTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using WTableCellProperties = DocumentFormat.OpenXml.Wordprocessing.TableCellProperties;
using WBold = DocumentFormat.OpenXml.Wordprocessing.Bold;
using WRunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;

namespace ClientDetailsApp
{
    internal class WordService
    {
        /// <summary>
        /// Reads all text from the first page of a docx file.
        /// Stops at an explicit page break or a paragraph marked with PageBreakBefore.
        /// </summary>
        public string ReadFirstPage(string filePath)
        {
            using var doc = WordprocessingDocument.Open(filePath, false);
            var body = doc.MainDocumentPart!.Document!.Body!;
            var sb = new StringBuilder();

            foreach (var element in body.ChildElements)
            {
                // Case 1: paragraph marked "page break before"
                if (element is DocumentFormat.OpenXml.Wordprocessing.Paragraph para)
                {
                    var pPr = para.ParagraphProperties;
                    if (pPr?.PageBreakBefore?.Val?.Value != false && pPr?.PageBreakBefore != null)
                        break;
                }

                // Case 2: explicit <w:br w:type="page"/>
                bool hasExplicitBreak = false;
                foreach (var run in element.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>())
                {
                    foreach (var child in run.ChildElements)
                    {
                        if (child is DocumentFormat.OpenXml.Wordprocessing.Break br &&
                            br.Type?.Value == DocumentFormat.OpenXml.Wordprocessing.BreakValues.Page)
                        {
                            hasExplicitBreak = true;
                            break;
                        }
                        if (child is DocumentFormat.OpenXml.Wordprocessing.Text t)
                            sb.Append(t.Text);
                    }
                    if (hasExplicitBreak) break;
                }

                if (hasExplicitBreak) break;

                if (!element.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().Any())
                    sb.AppendLine(element.InnerText);
                else
                    sb.AppendLine();
            }

            return sb.ToString().Trim();
        }

        public void OpenHeara(ClientDetails clientDetails)
        {
            string configPath = Path.Combine(AppContext.BaseDirectory, "appsettings.json");
            string configJson = File.ReadAllText(configPath);
            using var config = JsonDocument.Parse(configJson);
            string templatePath = config.RootElement.GetProperty("HearaTemplatePath").GetString()!;

            string tempPath = Path.Combine(Path.GetTempPath(), $"הערת_אזהרה_{DateTime.Now:yyyyMMdd_HHmmss}.dotx");
            File.Copy(templatePath, tempPath, overwrite: true);

            using (var wordDoc = WordprocessingDocument.Open(tempPath, true))
            {
                var body = wordDoc.MainDocumentPart!.Document!.Body!;
                var fieldMap = config.RootElement.GetProperty("FieldMap");
                foreach (var entry in fieldMap.EnumerateObject())
                    ReplaceText(body, entry.Name, ResolveField(clientDetails, entry.Value.GetString()!));
                wordDoc.MainDocumentPart.Document.Save();
            }

            System.Diagnostics.Process.Start(
                new System.Diagnostics.ProcessStartInfo(tempPath) { UseShellExecute = true });
        }

        public void OpenShatarMecher(ClientDetails clientDetails)
        {
            string configPath = Path.Combine(AppContext.BaseDirectory, "appsettings.json");
            string configJson = File.ReadAllText(configPath);
            using var config = JsonDocument.Parse(configJson);
            string templatePath = config.RootElement.GetProperty("ShatarMecherTemplatePath").GetString()!;
            string tempPath = Path.Combine(Path.GetTempPath(), $"שטר_מכר_{DateTime.Now:yyyyMMdd_HHmmss}.dotx");

            File.Copy(templatePath, tempPath, overwrite: true);

            using (var wordDoc = WordprocessingDocument.Open(tempPath, true))
            {
                var body = wordDoc.MainDocumentPart!.Document!.Body!;
                PopulateSellersTable(body, clientDetails.Sellers);
                PopulateBuyersTable(body, clientDetails.Buyers);
                PopulatePropertyTable(body, clientDetails.Property,"מס' הגוש");
                PopulateSignatureTable(body, clientDetails.Sellers, clientDetails.Buyers);
                wordDoc.MainDocumentPart.Document.Save();
            }

            System.Diagnostics.Process.Start(
                new System.Diagnostics.ProcessStartInfo(tempPath) { UseShellExecute = true });
        }

        private static void PopulateBuyersTable(WBody body, List<Person> buyers)
        {
            var table = body.ChildElements
                .SkipWhile(el => !el.InnerText.Contains("שם משפחה ושם פרטי"))
                .Skip(1)
                .SkipWhile(el => !el.InnerText.Contains("שם משפחה ושם פרטי"))
                .OfType<WTable>()
                .FirstOrDefault();
            if (table == null) return;

            var rows = table.Elements<WTableRow>().ToList();
            if (rows.Count < 2) return;

            var templateRow = (WTableRow)rows[1].CloneNode(true);

            foreach (var row in rows.Skip(1).ToList())
                row.Remove();

            string share = $"1/{buyers.Count}";

            foreach (var buyer in buyers)
            {
                var newRow = (WTableRow)templateRow.CloneNode(true);
                var cells = newRow.Elements<WTableCell>().ToList();

                string[] values = [buyer.FullName, "ת.ז.", buyer.Id, share];
                for (int i = 0; i < cells.Count && i < values.Length; i++)
                    FillCellFormField(cells[i], values[i]);

                table.AppendChild(newRow);
            }
        }

        private static void PopulateSignatureTable(WBody body, List<Person> sellers, List<Person> buyers)
        {
            var table = body.ChildElements
                .SkipWhile(el => !el.InnerText.Contains("חתימה"))
                .OfType<WTable>()
                .FirstOrDefault();
            if (table == null) return;

            var rows = table.Elements<WTableRow>().ToList();
            if (rows.Count < 2) return;

            var templateRow = (WTableRow)rows[2].CloneNode(true);

            foreach (var row in rows.Skip(2).ToList())
                row.Remove();

            int rowCount = Math.Max(sellers.Count, buyers.Count);
            for (int i = 0; i < rowCount; i++)
            {
                string sellerName = i < sellers.Count ? sellers[i].FullName : "";
                string buyerName  = i < buyers.Count  ? buyers[i].FullName  : "";

                var newRow = (WTableRow)templateRow.CloneNode(true);
                var cells = newRow.Elements<WTableCell>().ToList();

                string[] values = [sellerName, "", "", buyerName, ""];
                for (int j = 0; j < cells.Count && j < values.Length; j++)
                    FillCellFormField(cells[j], values[j]);

                table.AppendChild(newRow);
            }
        }

        private static void PopulatePropertyTable(WBody body, Property property, string tableIdentifier)
        {
            var table = body.ChildElements
                .SkipWhile(el => !el.InnerText.Contains(tableIdentifier))
                .OfType<WTable>()
                .FirstOrDefault();
            if (table == null) return;

            var firstDataRow = table.Elements<WTableRow>().Skip(5).FirstOrDefault();
            if (firstDataRow == null) return;

            var cells = firstDataRow.Elements<WTableCell>().ToList();

            string[] values = [property.Block, property.Parcel + "/" + property.SubParcel];
            for (int i = 0; i < cells.Count && i < values.Length; i++)
                FillCellFormField(cells[i], values[i]);
        }

        private static void PopulateSellersTable(WBody body, List<Person> sellers)
        {
            var table = body.ChildElements
                .SkipWhile(el => !el.InnerText.Contains("שם משפחה ושם פרטי"))
                .OfType<WTable>()
                .FirstOrDefault();
            if (table == null) return;

            var rows = table.Elements<WTableRow>().ToList();
            if (rows.Count < 2) return;

            // Clone the template data row before removing anything
            var templateRow = (WTableRow)rows[1].CloneNode(true);

            // Remove all data rows, keep only the header
            foreach (var row in rows.Skip(1).ToList())
                row.Remove();

            foreach (var seller in sellers)
            {
                var newRow = (WTableRow)templateRow.CloneNode(true);
                var cells = newRow.Elements<WTableCell>().ToList();

                string[] values = [seller.FullName, "ת.ז.", seller.Id];
                for (int i = 0; i < cells.Count && i < values.Length; i++)
                    FillCellFormField(cells[i], values[i]);

                table.AppendChild(newRow);
            }
        }

        // Fills the value run of a legacy form field (between "separate" and "end" FieldChar markers).
        // Falls back to replacing the first text run if no form field is found.
        private static string ResolveField(ClientDetails clientDetails, string fieldPath)
        {
            var parts = fieldPath.Split('.');
            if (parts.Length == 3 && int.TryParse(parts[1], out int index))
            {
                index--; // convert 1-based to 0-based
                var list = parts[0] == "Buyer" ? clientDetails.Buyers : clientDetails.Sellers;
                if (index < 0 || index >= list.Count) return "";
                var person = list[index];
                return parts[2] switch
                {
                    "FirstName" => person.FirstName,
                    "LastName"  => person.LastName,
                    "Id"        => person.Id,
                    "Share"     => person.Share,
                    "Address"   => person.Address,
                    _           => ""
                };
            }

            return fieldPath switch
            {
                "Property.Block"     => clientDetails.Property.Block,
                "Property.Parcel"    => clientDetails.Property.Parcel,
                "Property.SubParcel" => clientDetails.Property.SubParcel,
                "Property.Address"   => clientDetails.Property.PropAddress,
                _                    => ""
            };
        }

        private static void ReplaceText(WBody body, string search, string replacement)
        {
            var regex = new System.Text.RegularExpressions.Regex($@"\b{System.Text.RegularExpressions.Regex.Escape(search)}\b");

            foreach (var paragraph in body.Descendants<WParagraph>())
            {
                var runs = paragraph.Elements<WRun>()
                    .Where(r => r.GetFirstChild<WText>() != null)
                    .ToList();

                if (runs.Count == 0) continue;

                string fullText = string.Concat(runs.Select(r => r.GetFirstChild<WText>()!.Text));
                if (!regex.IsMatch(fullText)) continue;

                // Write replaced text into the first run, clear the rest
                var firstText = runs[0].GetFirstChild<WText>()!;
                firstText.Text = regex.Replace(fullText, replacement);
                firstText.Space = SpaceProcessingModeValues.Preserve;

                foreach (var run in runs.Skip(1))
                    run.GetFirstChild<WText>()!.Text = "";
            }
        }

        private static void FillCellFormField(WTableCell cell, string value)
        {
            var runs = cell.Descendants<WRun>().ToList();
            bool afterSeparate = false;

            foreach (var run in runs)
            {
                var fldChar = run.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.FieldChar>();
                if (fldChar != null)
                {
                    var type = fldChar.FieldCharType?.Value;
                    if (type == DocumentFormat.OpenXml.Wordprocessing.FieldCharValues.Separate)
                        afterSeparate = true;
                    else if (type == DocumentFormat.OpenXml.Wordprocessing.FieldCharValues.End)
                        afterSeparate = false;
                }
                else if (afterSeparate)
                {
                    var text = run.GetFirstChild<WText>();
                    if (text != null)
                    {
                        text.Text = value;
                        text.Space = SpaceProcessingModeValues.Preserve;
                        return;
                    }
                }
            }

            // Fallback: no form field structure — just set the first text run
            var fallback = cell.Descendants<WText>().FirstOrDefault();
            if (fallback != null) fallback.Text = value;
        }

        /// <summary>
        /// Creates a new docx at a temp path from a ClientDetails instance, then opens it.
        /// </summary>
        public void CreateAndOpen(ClientDetails clientDetails)
        {
            string path = Path.Combine(Path.GetTempPath(), $"עסקה_{DateTime.Now:yyyyMMdd_HHmmss}.docx");

            using (var wordDoc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new WDocument(new WBody());
                var body = mainPart.Document.Body!;

                AppendPartiesSection(body, clientDetails);
                AppendPropertySection(body, clientDetails.Property);
                AppendLawyersSection(body, clientDetails.Lawyers);

                mainPart.Document.Save();
            }

            System.Diagnostics.Process.Start(
                new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
        }

        private static void AppendPartiesSection(WBody body, ClientDetails clientDetails)
        {
            var rows = new List<string[]>();
            foreach (var p in clientDetails.Buyers)
                rows.Add(["קונה", p.FirstName, p.LastName, p.Id]);
            foreach (var p in clientDetails.Sellers)
                rows.Add(["מוכר", p.FirstName, p.LastName, p.Id]);

            if (rows.Count == 0) return;

            string[] headers = ["סוג", "שם פרטי", "שם משפחה", "תעודת זהות"];
            AppendTable(body, "קונים ומוכרים", headers, rows);
        }

        private static void AppendPropertySection(WBody body, Property property)
        {
            string[] headers = ["כתובת", "גוש", "חלקה", "תת חלקה"];
            var rows = new List<string[]> { new[] { property.PropAddress, property.Block, property.Parcel, property.SubParcel } };
            AppendTable(body, "נכס", headers, rows);
        }

        private static void AppendLawyersSection(WBody body, Lawyers lawyers)
        {
            string[] headers = ["תפקיד", "שם"];
            var rows = new List<string[]>
            {
                new[] { "עורך דין קונה", lawyers.BuyerLawyer },
                new[] { "עורך דין מוכר", lawyers.SellerLawyer }
            };
            AppendTable(body, "עורכי דין", headers, rows);
        }

        private static void AppendTable(WBody body, string title, string[] headers, List<string[]> rows)
        {
            body.AppendChild(new WParagraph(new WRun(
                new WRunProperties(new WBold()),
                new WText(title))));

            var table = new WTable();

            var headerRow = new WTableRow();
            foreach (string h in headers)
                headerRow.AppendChild(new WTableCell(
                    new WTableCellProperties(),
                    new WParagraph(new WRun(new WRunProperties(new WBold()), new WText(h)))));
            table.AppendChild(headerRow);

            foreach (string[] row in rows)
            {
                var tableRow = new WTableRow();
                foreach (string cell in row)
                    tableRow.AppendChild(new WTableCell(
                        new WTableCellProperties(),
                        new WParagraph(new WRun(new WText(cell)))));
                table.AppendChild(tableRow);
            }

            body.AppendChild(table);
            body.AppendChild(new WParagraph()); // spacer
        }
    }
}
