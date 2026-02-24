using System.Text.Json;
using System.Text.Json.Serialization;

namespace PFMakerLib;
internal class PFConfig
{
    public required string[] docxFiles { get; set; }
    public required Section[] sections { get; set; }

}

class Section
{
    public required string[] files { get; set; }
    public Vars? vars { get; set; }
    public Block[] blocks { get; set; } = [];
    public Replace[] replace { get; set; } = [];
}

class Vars
{
    [JsonExtensionData]
    public Dictionary<string, JsonElement>? Fields { get; set; }
}

class Block
{
    public bool? skip { get; set; }
    public string? var { get; set; }
    public string? mark { get; set; }
    public required string tag { get; set; }
    public string? postRE { get; set; }
    public string? preRE { get; set; }
    public required string before { get; set; }
    public required string after { get; set; }
    public string? text { get; set; }
}

class Replace
{
    public required string re { get; set; }
    public required string newText { get; set; }
}