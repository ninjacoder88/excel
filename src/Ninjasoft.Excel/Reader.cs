using System.Diagnostics;
using System.IO.Compression;
using System.Xml;
using System.Xml.Serialization;

namespace Ninjasoft.Excel;

// public class Reader
// {
//     public bool Read(string filePath)
//     {
//         if(string.IsNullOrWhiteSpace(filePath))
//             throw new Exception("filePath must have a value");

//         if(!File.Exists(filePath))
//             throw new Exception("filePath does not exist");

//         string? extension = Path.GetExtension(filePath);
//         if(string.IsNullOrWhiteSpace(extension))
//             throw new Exception("file extension must have a value");

//         if(extension != ".xlsx")
//             throw new Exception($"file extension must be xlsx; {extension}");

//         string fileName = Path.GetFileNameWithoutExtension(filePath);
        
//         DirectoryInfo outputDirectory = Directory.CreateDirectory($"/home/t/Downloads/{fileName}");

//         using(FileStream fs = File.OpenRead(filePath))
//         {
//             using(ZipArchive archive = new(fs))
//             {
//                 archive.ExtractToDirectory($"/home/t/Downloads/{fileName}");
//             }
//         }

        

//         return true;        
//     }
// }

public class ExcelFile
{
    public ExcelFile(string filePath)
    {
        if(string.IsNullOrWhiteSpace(filePath))
            throw new Exception("filePath must have a value");

        if(!File.Exists(filePath))
            throw new Exception($"file does not exist at path {filePath}");

        string? extension = Path.GetExtension(filePath);
        if(string.IsNullOrWhiteSpace(extension))
            throw new Exception("file extension must have a value");
        
        if(extension != ".xlsx")
            throw new Exception($"file extension must be xlsx; {extension}");

        Extension = extension;
        FilePath = filePath;
    }

    public string FilePath {get;}

    public string FileNameNoExt => Path.GetFileNameWithoutExtension(FilePath);

    public string Extension {get;}

    public void Read()
    {
        string outputDirectoryPath = $"/home/t/Downloads/{FileNameNoExt}";

        DirectoryInfo outputDirectory;
        if(!Directory.Exists(outputDirectoryPath))
        {
            Console.WriteLine("directory does not exist");
            outputDirectory = Directory.CreateDirectory(outputDirectoryPath);

            using(FileStream fs = File.OpenRead(FilePath))
            {
                using(ZipArchive archive = new(fs))
                {
                    archive.ExtractToDirectory(outputDirectoryPath);
                }
            }
        }
        else
            outputDirectory = new(outputDirectoryPath);
            

        DirectoryInfo? xlDirectory = new(Path.Combine(outputDirectory.FullName, "xl"));
        if(xlDirectory == null)
            throw new FileLoadException("xl directory is missing");

        DirectoryInfo? worksheetsDirectory = new(Path.Combine(xlDirectory.FullName, "worksheets"));
        if(worksheetsDirectory == null)
            throw new FileLoadException("worksheets directory is missing");

        FileInfo workbookFile = new(Path.Combine(xlDirectory.FullName, "workbook.xml"));
        if(workbookFile == null)
            throw new FileLoadException("workbook.xml is missing");

        FileInfo sheet1File = new(Path.Combine(worksheetsDirectory.FullName, "sheet1.xml"));
        if(sheet1File == null)
            throw new FileLoadException("sheet1.xml is missing");

        FileInfo sharedStringsFile = new(Path.Combine(xlDirectory.FullName, "sharedStrings.xml"));
        if(sharedStringsFile == null)
            throw new Exception("sharedStrings.xml is missing");


        XmlSerializer workbookSerializer = new(typeof(Workbook));
        
        object? workbookObj = workbookSerializer.Deserialize(new MyXmlTextReader(File.OpenRead(workbookFile.FullName)));
        if(workbookObj == null)
            throw new FileLoadException("failed to deserialize workbook file");

        Workbook? workbook = workbookObj as Workbook;
        if(workbook == null)
            throw new FileLoadException("failed to deserialize workbook to Workbook");
        

        XmlSerializer worksheetSerializer = new(typeof(Worksheet));

        object? worksheetObj = worksheetSerializer.Deserialize(new MyXmlTextReader(File.OpenRead(sheet1File.FullName)));
        if(worksheetObj == null)
            throw new FileLoadException("failed to deserialize worksheet file");

        Worksheet? worksheet = worksheetObj as Worksheet;
        if(worksheet == null)
            throw new FileLoadException("failed to deserialize worksheet to Worksheet");

        foreach(var x in worksheet.SheetData)
        {
            Console.WriteLine(x.R);
            foreach(var y in x.Columns)
            {
                Console.WriteLine($"{x.R} {y.Position}");
            }
        }


        XmlSerializer sharedStringSerializer = new(typeof(SharedStrings));

        object? sharedStringsObj = sharedStringSerializer.Deserialize(new MyXmlTextReader(File.OpenRead(sharedStringsFile.FullName)));
        if(sharedStringsObj == null)
            throw new FileLoadException("failed to deserialize shared strings file");

        SharedStrings? sharedStrings = sharedStringsObj as SharedStrings;
        if(sharedStrings == null)
            throw new FileLoadException("failed to deserialize sharedStrings to SharedStrings");

        foreach(var x in sharedStrings.Strings)
            Console.WriteLine(x.Text);
    }
}

public class MyXmlTextReader : XmlTextReader
{
    public MyXmlTextReader(Stream stream) : base(stream)
    { 
    }

    public override string NamespaceURI => string.Empty;
}

[XmlRoot("workbook", Namespace = "")]
public class Workbook
{
    [XmlElement("fileVersion")]
    public FileVersion FileVersion {get;set;}

    [XmlArray("sheets")]
    [XmlArrayItem("sheet")]
    public List<Sheet> Sheets {get;set;}
}

public class FileVersion
{
    [XmlAttribute("appName")]
    public string AppName {get;set;}
}

public class Sheet
{
    [XmlAttribute("name")]
    public string Name {get;set;}

    [XmlAttribute("sheetId")]
    public int SheetId {get;set;}
}



[XmlRoot("worksheet")]
public class Worksheet
{
    [XmlArray("sheetData")]
    [XmlArrayItem("row")]
    public List<Row> SheetData {get;set;}
}

public class Row
{
    [XmlElement("c")]
    public List<Column> Columns {get;set;}

    [XmlAttribute("r")]
    public int R {get;set;}
}

public class Column
{
    [XmlAttribute("r")]
    public string Position {get;set;}

    [XmlAttribute("t")]
    public string ContentType {get;set;}
}


[XmlRoot("sst")]
public class SharedStrings
{
    [XmlAttribute("count")]
    public int Count {get;set;}

    [XmlAttribute("uniqueCount")]
    public int UniqueCount {get;set;}

    [XmlElement("si")]
    public List<StringInfo> Strings {get;set;}
}

public class StringInfo
{
    [XmlElement("t")]
    public string Text {get;set;}
}