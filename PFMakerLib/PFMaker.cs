using Hjson;
using System.IO.Compression;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace PFMakerLib;

public class PFMaker
{
    private readonly string BaseTmpFolderPath = Path.GetTempPath();
    private readonly string tempFolder = "WordPrintFormMaker";
    private string TempFolderPath => Path.Combine([this.BaseTmpFolderPath, this.tempFolder]);


    public Result MakePrintForm(string cfgFilePath)
    {
        if (!File.Exists(cfgFilePath))
        {
            return Result.Failure("JSON file not found");
        }

        string? folderPath = Path.GetDirectoryName(cfgFilePath);

        if (folderPath == null)
        {
            return Result.Failure("Path of JSON file not found");
        }

        string jsonText = File.ReadAllText(cfgFilePath);

        PFConfig? config;

        try
        {
            var jsonValue = HjsonValue.Parse(jsonText);
            config = JsonSerializer.Deserialize<PFConfig>(jsonValue.ToString());
        }
        catch (Exception e)
        {
            return Result.Failure($"JSON parsing error: {e.Message}");
        }

        if (config == null)
        {
            return Result.Failure("Invalid JSON");
        }


        string docxFilePath;

        foreach (var docxFile in config.docxFiles)
        {
            docxFilePath = Path.Combine([folderPath, docxFile]);

            if (!File.Exists(docxFilePath))
            {
                return Result.Failure("File not found: " + docxFile);
            }

            var result = DocxToPrintForm(docxFilePath, config.sections);

            if (!result.IsSuccess)
            {
                return result;
            }
        }

        return Result.Success();
    }

    private Result DocxToPrintForm(string docxFilePath, Section[] sections)
    {
        UnzipDocx(docxFilePath);

        string sectionFilePath;

        foreach (var section in sections)
        {
            foreach (var sectonFile in section.files)
            {
                sectionFilePath = Path.Combine(this.TempFolderPath, sectonFile);

                ReplaceFileData(sectionFilePath, section);
            }
        }

        string? printFormFilePath = GetPrintFormFilePath(docxFilePath);

        if (printFormFilePath is null)
        {
            return Result.Failure("Target folder path not found");
        }

        ZipDocx(printFormFilePath);

        return Result.Success();
    }

    private void ReplaceFileData(string filePath, Section section)
    {
        string fileContent = File.ReadAllText(filePath);

        if (section.replace.Length > 0)
            fileContent = ReplaceRegExElems(fileContent, section.replace);

        if (section.blocks.Length > 0)
            fileContent = ReplaceBlocks(fileContent, section.blocks);

        if (section.vars is not null)
            fileContent = ReplaceVars(fileContent, section.vars);

        fileContent = RemoveMarks(fileContent);

        File.WriteAllText(filePath, fileContent);
    }

    private string ReplaceVars(string content, Vars textVars)
    {
        if (textVars.Fields is null)
            return content;

        foreach (var field in textVars.Fields)
        {
            content = content.Replace("[#" + field.Key + "#]", "<%=XmlAttrEncode(" + field.Value + ")%>");
        }

        return content;
    }

    private string ReplaceBlocks(string content, Block[] blocks)
    {
        foreach (var block in blocks)
        {
            if (block.skip ?? false)
                continue;

            string tag = block.tag;

            string centralText = "";

            string preRE = block.preRE ?? "";
            string postRE = block.postRE ?? "";

            if (block.var is not null)
                centralText = @"\[\#" + block.var + @"\#\]";
            else if (block.mark is not null)
                centralText = @"\#\#" + block.mark + @"\#\#";
            else if (block.text is not null)
                centralText = block.text;

            string pattern = "(" + preRE + @"<" + tag + @"(?:(?!(<\/" + tag + @">)).)*?" + centralText + @".*?<\/" + tag + @">" + postRE + ")";
            content = Regex.Replace(content, pattern, block.before + "$1" + block.after, RegexOptions.Singleline);
        }

        return content;
    }

    private string ReplaceRegExElems(string content, Replace[] reElems)
    {
        foreach (var reElem in reElems)
        {
            content = Regex.Replace(content, reElem.re, reElem.newText, RegexOptions.Singleline);
        }

        return content;
    }

    private string RemoveMarks(string content)
    {
        string pattern = @"\#\#.*?\#\#"; ;
        content = Regex.Replace(content, pattern, "", RegexOptions.Singleline);

        return content;
    }

    private void UnzipDocx(string docxFilePath)
    {
        if (Directory.Exists(this.TempFolderPath))
            Directory.Delete(this.TempFolderPath, true);

        ZipFile.ExtractToDirectory(docxFilePath, this.TempFolderPath);
    }

    private void ZipDocx(string docxFilePath)
    {
        if (File.Exists(docxFilePath))
            File.Delete(docxFilePath);

        ZipFile.CreateFromDirectory(this.TempFolderPath, docxFilePath);

        Directory.Delete(this.TempFolderPath, true);
    }

    private string? GetPrintFormFilePath(string docxFilePath)
    {
        string fileName = Path.GetFileNameWithoutExtension(docxFilePath);
        string fileExt = Path.GetExtension(docxFilePath);
        string? fileFolderPath = Path.GetDirectoryName(docxFilePath);

        if (fileFolderPath is null)
            return null;

        string pfFilePath = Path.Combine([
                fileFolderPath,
                fileName+"_PF"+fileExt
            ]);

        return pfFilePath;
    }

}
