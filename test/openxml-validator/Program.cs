using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

if (args.Length == 0)
{
    Console.Error.WriteLine("用法: MDDTT.OpenXmlValidator <文件.docx> [更多文件.docx ...]");
    return 2;
}

var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
var failed = false;

foreach (var input in args)
{
    var file = Path.GetFullPath(input);
    if (!File.Exists(file))
    {
        Console.Error.WriteLine($"失败 | Open XML Schema: 找不到文件 {file}");
        failed = true;
        continue;
    }

    try
    {
        using var document = WordprocessingDocument.Open(file, false);
        var errors = validator.Validate(document).ToList();

        if (errors.Count == 0)
        {
            Console.WriteLine($"通过 | Open XML Schema: {file}");
            continue;
        }

        failed = true;
        Console.Error.WriteLine($"失败 | Open XML Schema: {file}（{errors.Count} 项错误）");
        foreach (var error in errors)
        {
            Console.Error.WriteLine($"  [{error.Id ?? "Unknown"}] {error.Description}");
            if (error.Part?.Uri is not null)
                Console.Error.WriteLine($"    部件: {error.Part.Uri}");
            if (!string.IsNullOrWhiteSpace(error.Path?.XPath))
                Console.Error.WriteLine($"    路径: {error.Path.XPath}");
        }
    }
    catch (Exception exception)
    {
        failed = true;
        Console.Error.WriteLine($"失败 | Open XML Schema: {file}");
        Console.Error.WriteLine($"  无法打开或验证文档: {exception.Message}");
    }
}

return failed ? 1 : 0;
