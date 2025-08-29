using System.Text.Json.Serialization;

[JsonSerializable(typeof(Dictionary<string, List<Dictionary<string, object?>>>))]
[JsonSourceGenerationOptions(WriteIndented = true)]
internal partial class XltuiJsonContext : JsonSerializerContext
{
}
