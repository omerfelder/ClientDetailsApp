using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace ClientDetailsApp
{
    internal class OpenAiService
    {
        private readonly string _apiKey;

        private const string SystemPrompt =
            "תפקידך הוא מציאת כל הפרטים הרלוונטיים בהסכם.\n" +
            "בכל הסכם יכולים להיות מספר קונים ומספר מוכרים. לכל קונה ומוכר יש שם פרטי, שם משפחה, תעודת זהות וכתובת.\n" +
            "במידה והשם הפרטי מכיל שתי מילים ואין שם משפחה, פצל את השם הפרטי.\n" +
            "במידה וישנה רק כתובת אחת לקונים יש לשכפל אותה לכל הקונים. במידה וישנה רק כתובת אחת למוכרים יש לשכפל אותה לכל המוכרים.\n" +
            "במידה והחלקה מופיעה בפורמט של מספר/מספר אז יש לחלק את שני המספרים.לחלקה ותת חלקה\n" +
            "החזר אך ורק JSON תקני ללא טקסט נוסף, בתבנית הבאה:\n" +
            "{\n" +
            "  \"קונים\": [ { \"שם פרטי\": \"...\", \"שם משפחה\": \"...\", \"תעודת זהות\": \"...\", \"כתובת\": \"...\" } ],\n" +
            "  \"מוכרים\": [ { \"שם פרטי\": \"...\", \"שם משפחה\": \"...\", \"תעודת זהות\": \"...\", \"כתובת\": \"...\" } } ],\n" +
            "  \"נכס\": { \"כתובת\": \"...\", \"גוש\": \"...\", \"חלקה\": \"...\", \"תת חלקה\": \"...\" },\n" +
            "  \"עורכי דין\": { \"עורך דין קונה\": \"...\", \"עורך דין מוכר\": \"...\" }\n" +
            "}\n" +
            "במידה ולא נמצא שם משפחה ובשם הפרטי יש שתי מילים, העבר את המילה השנייה לשם המשפחה.";

        public OpenAiService(string apiKey)
        {
            _apiKey = apiKey;
        }

        public async Task<string> AnalyzeContractAsync(string text)
        {
            using HttpClient http = new HttpClient();
            http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);

            var body = new
            {
                model = "gpt-3.5-turbo",
                messages = new[]
                {
                    new { role = "system", content = SystemPrompt },
                    new { role = "user", content = text }
                },
                temperature = 0.7,
                max_tokens = 1024
            };

            string json = JsonSerializer.Serialize(body);
            using var content = new StringContent(json, Encoding.UTF8, "application/json");
            HttpResponseMessage res = await http.PostAsync("https://api.openai.com/v1/chat/completions", content);

            res.EnsureSuccessStatusCode();
            string resultJson = await res.Content.ReadAsStringAsync();

            using JsonDocument doc = JsonDocument.Parse(resultJson);
            var choice = doc.RootElement.GetProperty("choices")[0];
            return choice.GetProperty("message").GetProperty("content").GetString() ?? "(empty)";
        }
    }
}
