using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


string? inputDir = Console.ReadLine();
if (inputDir == "" || inputDir == null)
{
    Console.WriteLine("Default input file.");
    inputDir = @"C:\Users\wdmk4\source\git\PG\JP_NET_Lab_2";
};
string? outputFile = Console.ReadLine();
if (outputFile == "" || outputFile == null)
{
    Console.WriteLine("Default output file.");
    outputFile = @"C:\Users\wdmk4\source\git\PG\JP_NET_Lab_2\excel.xlsx";
}
int depth;
string? depthInput = Console.ReadLine();
if (!Int32.TryParse(depthInput, out depth))
{
    Console.WriteLine("Default depth.");
    depth = 0;
}

var ep = new ExcelPackage();
var data = GetAllDirsAndFiles(depth, new DirectoryInfo(inputDir), "...");


///         Dir Structure         ///
var dirStructureSheet = ep.Workbook.Worksheets.Add("Struktura katalogu");

foreach (var elem in data.Select((x, i) => new { Value = x, Index = i + 1 }))
{
    dirStructureSheet.Cells[elem.Index, 1].Value = elem.Value.name;
    if (elem.Value.isFile)
    {
        dirStructureSheet.Cells[elem.Index, 2].Value = (elem.Value.extension != "") ? elem.Value.extension : "no extension";
        dirStructureSheet.Cells[elem.Index, 3].Value = elem.Value.size;
    }
    dirStructureSheet.Cells[elem.Index, 4].Value = elem.Value.attributes;

    dirStructureSheet.Row(elem.Index).OutlineLevel = elem.Value.depth;
    dirStructureSheet.Row(elem.Index).Collapsed = true;
}
for (int i = 1; i < 4; i++) dirStructureSheet.Column(i).AutoFit();


///        Statistics         ///
var statisticsSheet = ep.Workbook.Worksheets.Add("Statystyki");

data.RemoveAll(s => s.isFile == false);
data.Sort((x, y) => y.size.CompareTo(x.size));

for (int i = 0; i < (data.Count < 10 ? data.Count : 10); i++)
{
    statisticsSheet.Cells[i + 1, 1].Value = data[i].name;
    statisticsSheet.Cells[i + 1, 2].Value = (data[i].extension != "") ? data[i].extension : "no extension";
    statisticsSheet.Cells[i + 1, 3].Value = data[i].size;
    statisticsSheet.Cells[i + 1, 4].Value = data[i].attributes;
}

List<string> addedFile = new List<string>();
for (int i = 0, position = 1; i < (data.Count < 10 ? data.Count : 10); i++)
{
    if (!addedFile.Contains(data[i].extension))
    {
        statisticsSheet.Cells[position, 5].Formula = "=B" + (i + 1);
        statisticsSheet.Cells[position, 6].Formula = "=COUNTIF(B1:B10,B" + (i + 1) + ")";
        statisticsSheet.Cells[position, 7].Formula = "=SUMIF(B1:B10,B" + (i + 1) + ",C1:C10)";
        position++;

        addedFile.Add(data[i].extension);
    }
}
for (int i = 1; i < (data.Count < 7 ? data.Count : 7); i++) statisticsSheet.Column(i).AutoFit();


///         Chart 1         ///
var amountChart = (statisticsSheet.Drawings.AddChart("Amount", eChartType.Pie3D) as ExcelPieChart);
if (amountChart == null)
    return;
amountChart.Title.Text = "% rozszerzeń ilościowo";amountChart.SetPosition(1, 0, 8, 0);
amountChart.SetSize(600, 300);amountChart.DataLabel.ShowCategory = true;
amountChart.DataLabel.ShowPercent = true;var not_used_result = (amountChart.Series.Add("F1:F" + addedFile.Count, "E1:E" + addedFile.Count) as ExcelPieChartSerie);///         Chart 2         ///
var percentChart = (statisticsSheet.Drawings.AddChart("Percent", eChartType.Pie3D) as ExcelPieChart);
if (percentChart == null)
    return;
percentChart.Title.Text = "% rozszerzeń wg rozmiaru";percentChart.SetPosition(17, 0, 8, 0);
percentChart.SetSize(600, 300);percentChart.DataLabel.ShowCategory = true;
percentChart.DataLabel.ShowPercent = true;not_used_result = (percentChart.Series.Add("G1:G" + addedFile.Count, "E1:E" + addedFile.Count) as ExcelPieChartSerie);
///         Save         ///
if (File.Exists(outputFile))
{
    try
    {
        var file = new FileInfo(outputFile);
        FileStream stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
        stream.Close();
    }
    catch (IOException)
    {
        Console.WriteLine("Cannot save to file!");
        return;
    }
}
File.WriteAllBytes(outputFile, ep.GetAsByteArray());


List<FolderElement> GetAllDirsAndFiles(int depth, DirectoryInfo path, string text = "", int currentDepth = 0)
{
    List<FolderElement> result = new List<FolderElement>();

    if (depth == currentDepth)
        return result;

    foreach (DirectoryInfo subDir in path.GetDirectories())
    {
        var temp = new FolderElement(
                name: text + "/" + subDir.Name,
                extension: "Dir",
                size: 0,
                attributes: subDir.Attributes.ToString(),
                isFile: false,
                depth: currentDepth
            );
        result.Add(temp);

        result.AddRange(GetAllDirsAndFiles(depth, subDir, text + "/" + subDir.Name, currentDepth + 1));
    }
    foreach (FileInfo file in path.GetFiles())
    {
        var temp = new FolderElement(
            name: text + "/" + file.Name,
            extension: file.Extension,
            size: file.Length,
            attributes: file.Attributes.ToString(),
            isFile: true,
            depth: currentDepth
        );
        result.Add(temp);
    }
    return result;
}

struct FolderElement
{
    public FolderElement(string name, string extension, long size, string attributes, bool isFile, int depth)
    {
        this.name = name;
        this.extension = extension;
        this.size = size;
        this.attributes = attributes;
        this.isFile = isFile;
        this.depth = depth;
    }
    public string name;
    public string extension;
    public long size;
    public string attributes;
    public bool isFile;
    public int depth;
};