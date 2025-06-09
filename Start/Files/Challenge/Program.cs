const string folder = "FileCollection";
const string resultsfile = "results.txt";

long XLCount = 0, DOCCount = 0, PPTCount = 0;
long XLSize = 0, DOCSize = 0, PPTSize = 0;
long totalfiles = 0;
long totalsize = 0;

bool IsOfficeFile(string filename) {
    if (filename.EndsWith(".xlsx") || filename.EndsWith(".docx") || filename.EndsWith(".pptx"))
        return true;
    return false;
}

DirectoryInfo di = new DirectoryInfo(folder);

foreach (FileInfo fi in di.EnumerateFiles())
{
    if (IsOfficeFile(fi.Name))
    {
        totalfiles++;
        totalsize += fi.Length;
        if (fi.Name.EndsWith(".xlsx"))
        {
            XLCount++;
            XLSize += fi.Length;
        }
        if (fi.Name.EndsWith(".docx"))
        {
            DOCCount++;
            DOCSize += fi.Length;
        }
        if (fi.Name.EndsWith(".pptx"))
        {
            PPTCount++;
            PPTSize += fi.Length;
        }
    }
}

using (StreamWriter sw = File.CreateText(resultsfile))
{
    sw.WriteLine("---- Results: ----");
    sw.WriteLine($"Total Files: {totalfiles}");
    sw.WriteLine($"Excel Files: {XLCount}");
    sw.WriteLine($"Word Files: {DOCCount}");
    sw.WriteLine($"PowerPoint Files: {PPTCount}");
    sw.WriteLine("----");
    sw.WriteLine($"Total Size: {totalsize:N0}");
    sw.WriteLine($"Excel Size: {XLSize:N0}");
    sw.WriteLine($"World Size: {DOCSize:N0}");
    sw.WriteLine($"PowerPoint Size: {PPTSize:N0}");
}