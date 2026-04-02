using System.Diagnostics;
using System.Text;
using System.Text.Json;

namespace ClientDetailsApp
{
    public partial class ClientDetailsForm : Form
    {
        private Button buttonUpload;
        private Button buttonTest;
        private Button buttonShatarMecher;
        private Button buttonCopyParties;
        private Button buttonCopyProperty;
        private Button buttonCopyLawyers;
        private Label labelParties;
        private Label labelProperty;
        private Label labelLawyers;
        private DataGridView dataGridParties;
        private DataGridView dataGridProperty;
        private DataGridView dataGridLawyers;

        private readonly WordService _wordService = new();
        private ClientDetails? _clientDetails;

        private const string TestJSON = "{\"קונים\":[{\"שם פרטי\":\"קונה\",\"שם משפחה\":\"שלישי\",\"תעודת זהות\":\"123456789\"},{\"שם פרטי\":\"סטפני שרה\",\"שם משפחה\":\"לוי אהרונוביץ\",\"תעודת זהות\":\"305663544\"},{\"שם פרטי\":\"גיל\",\"שם משפחה\":\"לוי\",\"תעודת זהות\":\"301300885\"}],\"מוכרים\":[{\"שם פרטי\":\"מרק\",\"שם משפחה\":\"גרינשפון\",\"תעודת זהות\":\"303466296\"},{\"שם פרטי\":\"גבאי\",\"שם משפחה\":\"גרינשפון בת-אל\",\"תעודת זהות\":\"305772154\"}],\"נכס\":{\"כתובת\":\"רחוב סיטקוב 2, רחובות\",\"גוש\":\"3799\",\"חלקה\":\"932\",\"תת חלקה\":\"39\"},\"עורכי דין\":{\"עורך דין קונה\":\"משה לוי\",\"עורך דין מוכר\":\"משה כהן\"}}";

        public ClientDetailsForm()
        {
            InitializeComponent();
            BuildUi();
        }

        private void BuildUi()
        {
            this.Text = "העתקת פרטי עסקה";
            this.ClientSize = new Size(815, 512);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.RightToLeft = RightToLeft.Yes;
            this.RightToLeftLayout = true;


            buttonUpload = new Button { Text = "העלה הסכם מכר", Left = 10, Top = 15, Width = 160, Height = 30 };
            buttonUpload.Click += async (s, e) => await ButtonUpload_Click();

            buttonTest = new Button { Text = "טען נתוני בדיקה", Left = 180, Top = 15, Width = 160, Height = 30 };
            buttonTest.Click += (s, e) => LoadTestData();

            buttonShatarMecher = new Button { Text = "צור שטר מכר", Left = 350, Top = 15, Width = 160, Height = 30 };
            buttonShatarMecher.Click += (s, e) => OpenShatarMecherDoc();

            labelParties = new Label { Text = "קונים ומוכרים:", Left = -10, Top = 60, Height = 20, TextAlign = ContentAlignment.TopRight, RightToLeft = RightToLeft.Yes };
            buttonCopyParties = MakeCopyButton(top: 50);
            buttonCopyParties.Click += (s, e) => CopyGridToCsv(dataGridParties);
            dataGridParties = MakeGrid(top: 80, height: 150);

            labelProperty = new Label { Text = "נכס:", Left = -58, Top = 255, Height = 20, TextAlign = ContentAlignment.TopRight, RightToLeft = RightToLeft.Yes };
            buttonCopyProperty = MakeCopyButton(top: 245);
            buttonCopyProperty.Click += (s, e) => CopyGridToCsv(dataGridProperty);
            dataGridProperty = MakeGrid(top: 275, height: 72);

            labelLawyers = new Label { Text = "עורכי דין:", Left = -38, Top = 372, Height = 20, TextAlign = ContentAlignment.TopRight, RightToLeft = RightToLeft.Yes };
            buttonCopyLawyers = MakeCopyButton(top: 362);
            buttonCopyLawyers.Click += (s, e) => CopyGridToCsv(dataGridLawyers);
            dataGridLawyers = MakeGrid(top: 392, height: 100);

            this.Controls.AddRange(new Control[]
            {
                buttonUpload, buttonTest, buttonShatarMecher,
                dataGridParties, labelParties, buttonCopyParties,
                dataGridProperty, labelProperty, buttonCopyProperty,
                dataGridLawyers, labelLawyers, buttonCopyLawyers
            });
        }

        // Left=735 places the button at the visual left edge of the grid (815 - 10 - 70 = 735) in RTL layout
        private static Button MakeCopyButton(int top) => new Button
        {
            Text = "העתק", Left = 735, Top = top, Width = 70, Height = 25
        };

        private static DataGridView MakeGrid(int top, int height) => new DataGridView
        {
            Left = 10, Top = top, Width = 795, Height = height,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            ColumnHeadersVisible = true,
            RowHeadersVisible = false,
            ScrollBars = ScrollBars.Both,
            RightToLeft = RightToLeft.Yes
        };

        private static void CopyGridToCsv(DataGridView grid)
        {
            if (grid.Columns.Count == 0) return;
            var sb = new StringBuilder();
            foreach (DataGridViewColumn col in grid.Columns)
                sb.Append(col.HeaderText).Append(',');
            sb.Length--;
            sb.AppendLine();
            foreach (DataGridViewRow row in grid.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                    sb.Append(cell.Value?.ToString() ?? "").Append(',');
                sb.Length--;
                sb.AppendLine();
            }
            Clipboard.SetText(sb.ToString());
        }

        private async Task ButtonUpload_Click()
        {
            using var dlg = new OpenFileDialog { Filter = "Word Documents|*.docx" };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            string path = dlg.FileName;
            if (!path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("יש לבחור קובץ מסוג docx בלבד.", "סוג קובץ שגוי", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY") ?? "";
            if (string.IsNullOrWhiteSpace(apiKey))
            {
                MessageBox.Show("OpenAI API key not found. Set OPENAI_API_KEY environment variable.", "Missing API Key", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            buttonUpload.Enabled = false;
            ClearAllGrids();

            try
            {
                string pageText = _wordService.ReadFirstPage(path);
                var openAi = new OpenAiService(apiKey);
                string response = await openAi.AnalyzeContractAsync(pageText[..Math.Min(2000, pageText.Length)]);
                
                _clientDetails = ClientDetails.Parse(response);
                PopulateGrids(response);
            }
            catch (Exception ex)
            {
                dataGridParties.Columns.Add("err", "שגיאה");
                dataGridParties.Rows.Add(ex.Message);
            }
            finally
            {
                buttonUpload.Enabled = true;
            }
        }

        private void OpenShatarMecherDoc()
        {
            if (_clientDetails == null)
            {
                MessageBox.Show("העתק את הנתונים לפני יצירת המסמך", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _wordService.OpenShatarMecher(_clientDetails);
        }

        private void LoadTestData()
        {
            _clientDetails = ClientDetails.Parse(TestJSON);
            PopulateGrids(TestJSON);
        }


        private void ClearAllGrids()
        {
            foreach (DataGridView g in new[] { dataGridParties, dataGridProperty, dataGridLawyers })
            {
                g.Rows.Clear();
                g.Columns.Clear();
            }
        }

        private void PopulateGrids(string response)
        {
            Debug.WriteLine(response);
            ClearAllGrids();

            using JsonDocument doc = JsonDocument.Parse(response);
            JsonElement root = doc.RootElement;

            PopulatePartiesGrid(root);

            if (root.TryGetProperty("נכס", out JsonElement נכס) && נכס.ValueKind == JsonValueKind.Object)
                PopulateObjectAsRow(dataGridProperty, נכס);

            if (root.TryGetProperty("עורכי דין", out JsonElement lawyers))
            {
                if (lawyers.ValueKind == JsonValueKind.Object)
                    PopulateKeyValueRows(dataGridLawyers, lawyers);
                else if (lawyers.ValueKind == JsonValueKind.Array)
                    PopulateArrayOfObjects(dataGridLawyers, lawyers);
            }
        }

        private void PopulatePartiesGrid(JsonElement root)
        {
            var people = new List<(string role, JsonElement obj)>();

            if (root.TryGetProperty("קונים", out JsonElement buyers) && buyers.ValueKind == JsonValueKind.Array)
                foreach (JsonElement p in buyers.EnumerateArray())
                    people.Add(("קונה", p));

            if (root.TryGetProperty("מוכרים", out JsonElement sellers) && sellers.ValueKind == JsonValueKind.Array)
                foreach (JsonElement p in sellers.EnumerateArray())
                    people.Add(("מוכר", p));

            if (people.Count == 0) return;

            var keys = new List<string> { "סוג" };
            foreach ((_, JsonElement obj) in people)
                if (obj.ValueKind == JsonValueKind.Object)
                    foreach (JsonProperty prop in obj.EnumerateObject())
                        if (!keys.Contains(prop.Name)) keys.Add(prop.Name);

            foreach (string key in keys)
                dataGridParties.Columns.Add(key, key);

            foreach ((string role, JsonElement obj) in people)
            {
                if (obj.ValueKind != JsonValueKind.Object) continue;
                var cells = new string[keys.Count];
                cells[0] = role;
                foreach (JsonProperty prop in obj.EnumerateObject())
                {
                    int idx = keys.IndexOf(prop.Name);
                    if (idx >= 0) cells[idx] = prop.Value.ToString();
                }
                dataGridParties.Rows.Add(cells);
            }
        }

        private static void PopulateObjectAsRow(DataGridView grid, JsonElement obj)
        {
            var keys = new List<string>();
            foreach (JsonProperty prop in obj.EnumerateObject()) keys.Add(prop.Name);
            foreach (string key in keys) grid.Columns.Add(key, key);
            var cells = new string[keys.Count];
            int i = 0;
            foreach (JsonProperty prop in obj.EnumerateObject()) cells[i++] = prop.Value.ToString();
            grid.Rows.Add(cells);
        }

        private static void PopulateKeyValueRows(DataGridView grid, JsonElement obj)
        {
            grid.Columns.Add("תפקיד", "תפקיד");
            grid.Columns.Add("שם", "שם");
            foreach (JsonProperty prop in obj.EnumerateObject())
                grid.Rows.Add(prop.Name, prop.Value.ToString());
        }

        private static void PopulateArrayOfObjects(DataGridView grid, JsonElement array)
        {
            var keys = new List<string>();
            foreach (JsonElement item in array.EnumerateArray())
                if (item.ValueKind == JsonValueKind.Object)
                    foreach (JsonProperty prop in item.EnumerateObject())
                        if (!keys.Contains(prop.Name)) keys.Add(prop.Name);

            foreach (string key in keys) grid.Columns.Add(key, key);

            foreach (JsonElement item in array.EnumerateArray())
            {
                if (item.ValueKind != JsonValueKind.Object) continue;
                var cells = new string[keys.Count];
                foreach (JsonProperty prop in item.EnumerateObject())
                {
                    int idx = keys.IndexOf(prop.Name);
                    if (idx >= 0) cells[idx] = prop.Value.ToString();
                }
                grid.Rows.Add(cells);
            }
        }
    }
}
