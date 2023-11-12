using LovePdf.Core;
using LovePdf.Model.Task;

namespace VocabularyExtractor.LovePdf;

/// <summary>
///     Convert PDF To Office Documents
/// </summary>
public class PdfToOfficeTask : LovePdfTask
{
    /// <inheritdoc />
    public override string ToolName => "pdfoffice";

    /// <summary>
    ///     Process the task
    /// </summary>
    /// <returns></returns>
    public ExecuteTaskResponse Process()
    {
        var parameters = new PdfToOfficeParams("xlsx");

        return base.Process(parameters);
    }

    /// <summary>
    ///     Process the task
    /// </summary>
    /// <param name="parameters"></param>
    /// <returns></returns>
    public ExecuteTaskResponse Process(PdfToOfficeParams parameters)
    {
        parameters ??= new PdfToOfficeParams("xlsx");

        return base.Process(parameters);
    }
}