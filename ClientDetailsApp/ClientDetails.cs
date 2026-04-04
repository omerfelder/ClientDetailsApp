using System.Text.Json;
using System.Text.Json.Serialization;

namespace ClientDetailsApp
{
    internal class Person
    {
        [JsonPropertyName("שם פרטי")]
        public string FirstName { get; set; } = "";

        [JsonPropertyName("שם משפחה")]
        public string LastName { get; set; } = "";

        [JsonPropertyName("תעודת זהות")]
        public string Id { get; set; } = "";

        public string FullName => $"{FirstName} {LastName}".Trim();
        public string Share { get; set; } = "";
    }

    internal class Property
    {
        [JsonPropertyName("כתובת")]
        public string Address { get; set; } = "";

        [JsonPropertyName("גוש")]
        public string Block { get; set; } = "";

        [JsonPropertyName("חלקה")]
        public string Parcel { get; set; } = "";

        [JsonPropertyName("תת חלקה")]
        public string SubParcel { get; set; } = "";
    }

    internal class Lawyers
    {
        [JsonPropertyName("עורך דין קונה")]
        public string BuyerLawyer { get; set; } = "";

        [JsonPropertyName("עורך דין מוכר")]
        public string SellerLawyer { get; set; } = "";
    }

    internal class ClientDetails
    {
        [JsonPropertyName("קונים")]
        public List<Person> Buyers { get; set; } = [];

        [JsonPropertyName("מוכרים")]
        public List<Person> Sellers { get; set; } = [];

        [JsonPropertyName("נכס")]
        public Property Property { get; set; } = new();

        [JsonPropertyName("עורכי דין")]
        public Lawyers Lawyers { get; set; } = new();

        public static ClientDetails Parse(string json)
        {
            var result = JsonSerializer.Deserialize<ClientDetails>(json)
                ?? throw new InvalidOperationException("Failed to parse client details from JSON.");

            string buyerShare  = $"1/{result.Buyers.Count}";
            string sellerShare = $"1/{result.Sellers.Count}";

            foreach (var buyer  in result.Buyers)  buyer.Share  = buyerShare;
            foreach (var seller in result.Sellers) seller.Share = sellerShare;

            return result;
        }
    }
}
