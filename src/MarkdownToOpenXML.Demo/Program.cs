using System.Diagnostics;

using MarkdownToOpenXML;

// Define input and output file names
const string demoInputFile = "demo.txt";
const string demoOutputFile = "demo.docx";

// Define paths to input and output files
var basePath = Path.Combine(Path.GetDirectoryName(Environment.ProcessPath) ?? string.Empty);
var demoInputPath = Path.Combine(basePath, demoInputFile);
var demoOutputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), demoOutputFile);

// Read markdown from input file
Console.WriteLine($"Loading {demoInputPath}...");
var markdown = File.ReadAllText(demoInputPath);

// Process markdown and write to output file
Console.WriteLine($"Writing to {demoOutputPath}...");
MarkdownToOpenXml.CreateDocX(markdown, demoOutputPath);

// Open the DOCX file using the default application
Console.WriteLine("Opening the DOCX file...");

var openProcess = System.Runtime.InteropServices.RuntimeInformation.OSDescription switch
{
    { } os when os.Contains("Windows") => "explorer.exe",
    { } os when os.Contains("Linux") => "xdg-open",
    { } os when os.Contains("Darwin") => "open",
    _ => null
};

if (openProcess != null)
{
    Process.Start(new ProcessStartInfo
    {
        FileName = openProcess,
        Arguments = demoOutputPath,
        UseShellExecute = true
    });
}
else
{
    Console.WriteLine("Unable to open the DOCX file.");
}