using LovePdf.Model.TaskParams;
using Newtonsoft.Json;

namespace VocabularyExtractor.LovePdf;

/// <summary>
/// Office To Pdf Params
/// </summary>
public class PdfToOfficeParams : BaseParams
{
    public PdfToOfficeParams(
        string convertTo)
    {
        ConvertTo = convertTo;
    }

    /// <summary>
    ///     Detailed
    /// </summary>
    [JsonProperty("convert_to")]
    public string ConvertTo { get; set; }
}