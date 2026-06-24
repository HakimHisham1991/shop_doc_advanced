using System.Globalization;
using System.Text;
using ClosedXML.Excel;

namespace ShopDoc;

internal static class Program
{
    private const string FinishTypeHeader =
        "Finish Type \n(Finish / Controlled Roughing / Free Roughing)";

    private static readonly string[] MaterialList =
        ["Aluminium", "Titanium", "Steel", "Bronze"];

    private static readonly string[] ProcessSpecsList =
        ["AIPI 03-11-001", "80-T-30-4010"];

    private static readonly string[] SurfaceTypeList =
        ["Finish", "Controlled Roughing", "Free Roughing"];

    private static readonly string[] MillingTypeList =
        ["End Milling", "Face Milling", "Drilling", "Reaming", "Turning"];

    private static readonly string[] ToolTypeList =
        ["Carbide", "HSS", "PCD"];

    private static readonly string[] MachiningTypeList =
        ["Conventional", "HSM"];

    public static int Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        try
        {
            Console.OutputEncoding = Encoding.UTF8;
        }
        catch (IOException)
        {
            // NX/Tcl exec often has no console handle — stdout may still be a pipe.
        }

        if (args.Length < 1)
        {
            Log("ERROR", "Script", "Usage: convert_csv_to_xlsx.exe input.csv [output.xlsx]");
            return 1;
        }

        var csvFile = Path.GetFullPath(args[0]);
        var xlsxFile = args.Length > 1
            ? Path.GetFullPath(args[1])
            : Path.Combine(
                Path.GetDirectoryName(csvFile) ?? ".",
                Path.GetFileNameWithoutExtension(csvFile) + ".xlsx");

        xlsxFile = ResolveUniqueXlsxPath(xlsxFile);

        Log("INFO", "Script", "convert_csv_to_xlsx.exe started (ClosedXML, no Excel COM)");
        Log("INFO", "Paths", "CSV input: " + csvFile);
        Log("INFO", "Paths", "XLSX output: " + xlsxFile);

        if (!File.Exists(csvFile))
        {
            Log("ERROR", "CSV", "Input CSV file not found: " + csvFile);
            return 1;
        }

        var outputDir = Path.GetDirectoryName(xlsxFile);
        if (string.IsNullOrEmpty(outputDir))
        {
            Log("ERROR", "Paths", "Invalid output path: " + xlsxFile);
            return 1;
        }

        try
        {
            Directory.CreateDirectory(outputDir);
            var probe = Path.Combine(outputDir, "_shopdoc_exe_probe.tmp");
            File.WriteAllText(probe, "probe");
            File.Delete(probe);
            Log("INFO", "Permission", "Output folder is writable: " + outputDir);
        }
        catch (Exception ex)
        {
            Log("ERROR", "Permission",
                "Cannot write to output folder (permission denied, read-only path, or policy block): "
                + outputDir + " - " + ex.Message);
            return 1;
        }

        List<string[]> rows;
        try
        {
            rows = CsvReader.ReadAllRows(csvFile);
        }
        catch (Exception ex)
        {
            Log("ERROR", "CSV", "Cannot read CSV: " + csvFile + " - " + ex.Message);
            return 1;
        }

        if (rows.Count == 0)
        {
            Log("ERROR", "CSV", "CSV file is empty: " + csvFile);
            return 1;
        }

        Log("INFO", "CSV", "CSV has " + rows.Count + " rows");

        try
        {
            using var workbook = new XLWorkbook();
            ApplyModernOfficeTheme(workbook.Theme);
            var worksheet = workbook.Worksheets.Add("Sheet1");

            var lastRow = rows.Count;
            var lastCol = rows.Max(r => r.Length);

            for (var r = 0; r < rows.Count; r++)
            {
                var row = rows[r];
                for (var c = 0; c < row.Length; c++)
                {
                    worksheet.Cell(r + 1, c + 1).Value = row[c];
                }
            }

            var tableRange = worksheet.Range(1, 1, lastRow, lastCol);
            tableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            tableRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            try
            {
                var table = tableRange.CreateTable();
                // Excel UI label: "Dark Teal, Table Style Medium 9" (TableStyleMedium9 + modern Office theme)
                table.Theme = XLTableTheme.TableStyleMedium9;
                table.ShowAutoFilter = true;
                Log("INFO", "Format", "Applied table style: Dark Teal, Table Style Medium 9");
            }
            catch (Exception ex)
            {
                Log("WARN", "Format", "Could not create Excel table (continuing without table style) - " + ex.Message);
            }

            worksheet.Columns(1, lastCol).AdjustToContents();
            worksheet.Row(1).Style.Alignment.WrapText = true;
            worksheet.Row(1).AdjustToContents();

            AddListValidation(worksheet, FindColumn(worksheet, "Material Type"), lastRow, MaterialList);
            AddListValidation(worksheet, FindColumn(worksheet, "Process Specs"), lastRow, ProcessSpecsList);
            AddListValidation(worksheet, FindFinishTypeColumn(worksheet), lastRow, SurfaceTypeList);
            AddListValidation(worksheet, FindColumn(worksheet, "Cutter Type"), lastRow, MillingTypeList);
            AddListValidation(worksheet, FindColumn(worksheet, "Tool Type (Carbide/HSS/PCD)"), lastRow, ToolTypeList);
            AddListValidation(worksheet, FindColumn(worksheet, "Machining Type (Conventional/HSM)"), lastRow, MachiningTypeList);

            AddBlankConditionalFormatting(worksheet, FindColumn(worksheet, "Material Type"), lastRow);
            AddBlankConditionalFormatting(worksheet, FindColumn(worksheet, "Process Specs"), lastRow);
            AddBlankConditionalFormatting(worksheet, FindColumn(worksheet, "Cutter Type"), lastRow);
            AddBlankConditionalFormatting(worksheet, FindColumn(worksheet, "Tool Type (Carbide/HSS/PCD)"), lastRow);
            AddBlankConditionalFormatting(worksheet, FindFinishTypeColumn(worksheet), lastRow);
            AddBlankConditionalFormatting(worksheet, FindColumn(worksheet, "Machining Type (Conventional/HSM)"), lastRow);

            worksheet.SheetView.FreezeRows(1);
            worksheet.SheetView.FreezeColumns(1);
            Log("INFO", "Format", "Applied freeze panes");

            workbook.SaveAs(xlsxFile);
            Log("INFO", "XLSX", "Saved XLSX successfully: " + xlsxFile);
            Console.WriteLine("SUCCESS: " + xlsxFile);
            return 0;
        }
        catch (Exception ex)
        {
            Log("ERROR", "XLSX", "Cannot save XLSX file: " + xlsxFile + " - " + ex.Message);
            Log("ERROR", "Permission",
                "Save failed - folder may be read-only, file locked, blocked by OneDrive/sync, or antivirus");
            return 1;
        }
    }

    private static string ResolveUniqueXlsxPath(string requestedPath)
    {
        if (!File.Exists(requestedPath))
            return requestedPath;

        var dir = Path.GetDirectoryName(requestedPath) ?? ".";
        var baseName = Path.GetFileNameWithoutExtension(requestedPath);
        var ext = Path.GetExtension(requestedPath);
        if (string.IsNullOrEmpty(ext))
            ext = ".xlsx";

        Log("INFO", "XLSX", "Target XLSX already exists (will not overwrite): " + requestedPath);

        for (var i = 1; i < 10_000; i++)
        {
            var candidate = Path.Combine(dir, baseName + "_" + i + ext);
            if (!File.Exists(candidate))
            {
                Log("INFO", "XLSX", "Using alternate output filename: " + candidate);
                return candidate;
            }
        }

        throw new IOException("Could not find an available XLSX filename (tried _1 through _9999).");
    }

    private static void ApplyModernOfficeTheme(IXLTheme theme)
    {
        // Match Excel 365 default theme accents so TableStyleMedium9 renders as Dark Teal Medium 9.
        theme.Text2 = XLColor.FromHtml("#0E2841");
        theme.Background2 = XLColor.FromHtml("#E8E8E8");
        theme.Accent1 = XLColor.FromHtml("#156082");
        theme.Accent2 = XLColor.FromHtml("#E97132");
        theme.Accent3 = XLColor.FromHtml("#196B24");
        theme.Accent4 = XLColor.FromHtml("#0F9ED5");
        theme.Accent5 = XLColor.FromHtml("#A02B93");
        theme.Accent6 = XLColor.FromHtml("#4EA72E");
    }

    private static void Log(string level, string category, string message) =>
        Console.WriteLine($"SHOPDOC_LOG|{level}|{category}|{message}");

    private static int FindColumn(IXLWorksheet ws, string colName)
    {
        var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
        for (var col = 1; col <= lastCol; col++)
        {
            var text = ws.Cell(1, col).GetString().Trim();
            if (string.Equals(text, colName, StringComparison.Ordinal))
                return col;
        }

        return 0;
    }

    private static int FindFinishTypeColumn(IXLWorksheet ws)
    {
        var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
        var expected = NormalizeHeaderText(FinishTypeHeader);

        for (var col = 1; col <= lastCol; col++)
        {
            if (NormalizeHeaderText(ws.Cell(1, col).GetString()) == expected)
                return col;
        }

        for (var col = 1; col <= lastCol; col++)
        {
            var v = NormalizeHeaderText(ws.Cell(1, col).GetString());
            if (v.StartsWith("Finish Type", StringComparison.OrdinalIgnoreCase)
                && v.Contains("(Finish / Controlled Roughing / Free Roughing)", StringComparison.OrdinalIgnoreCase))
            {
                return col;
            }
        }

        return 0;
    }

    private static string NormalizeHeaderText(string value) =>
        value.Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n');

    private static string BuildInlineListSource(string[] values) =>
        string.Join(",", values);

    private static void AddBlankConditionalFormatting(IXLWorksheet ws, int colNum, int lastRow)
    {
        if (colNum <= 0 || lastRow <= 1)
            return;

        try
        {
            var dataRange = ws.Range(2, colNum, lastRow, colNum);
            ApplyBlankCellHighlight(dataRange.AddConditionalFormat().WhenIsBlank());
            ApplyBlankCellHighlight(dataRange.AddConditionalFormat().WhenEquals(string.Empty));
            Log("INFO", "Format", "Added blank conditional formatting to column " + colNum);
        }
        catch (Exception ex)
        {
            Log("WARN", "Format", "Could not add blank conditional formatting to column " + colNum + " - " + ex.Message);
        }
    }

    private static void ApplyBlankCellHighlight(IXLStyle style)
    {
        style.Fill.SetBackgroundColor(XLColor.Yellow);
        style.Fill.PatternType = XLFillPatternValues.Solid;
        style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
        style.Border.SetOutsideBorderColor(XLColor.Red);
    }

    private static void AddListValidation(IXLWorksheet ws, int colNum, int lastRow, string[] values)
    {
        if (colNum <= 0 || lastRow <= 1 || values.Length == 0)
            return;

        try
        {
            var dataRange = ws.Range(2, colNum, lastRow, colNum);
            var validation = dataRange.CreateDataValidation();
            // Plain comma-separated list in Source box (same as Excel COM Validation.Add xlValidateList)
            validation.List("\"" + BuildInlineListSource(values) + "\"", true);
            validation.InCellDropdown = true;
            validation.IgnoreBlanks = true;
            validation.ShowInputMessage = false;
            validation.ShowErrorMessage = true;
            validation.ErrorStyle = XLErrorStyle.Stop;
            validation.ErrorTitle = "Invalid value";
            validation.ErrorMessage = "Value must match one of the listed items.";
            Log("INFO", "Format", "Added validation to column " + colNum);
        }
        catch (Exception ex)
        {
            Log("WARN", "Format", "Could not add validation to column " + colNum + " - " + ex.Message);
        }
    }
}

internal static class CsvReader
{
    public static List<string[]> ReadAllRows(string path)
    {
        using var stream = File.OpenRead(path);
        using var reader = new StreamReader(stream, DetectEncoding(path), detectEncodingFromByteOrderMarks: true);
        var rows = new List<string[]>();
        var fields = new List<string>();
        var field = new StringBuilder();
        var inQuotes = false;

        while (true)
        {
            var ch = reader.Read();
            if (ch == -1)
            {
                if (field.Length > 0 || fields.Count > 0)
                    fields.Add(field.ToString());
                if (fields.Count > 0)
                    rows.Add(fields.ToArray());
                break;
            }

            var c = (char)ch;
            if (inQuotes)
            {
                if (c == '"')
                {
                    var next = reader.Peek();
                    if (next == '"')
                    {
                        reader.Read();
                        field.Append('"');
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    field.Append(c);
                }
            }
            else if (c == '"')
            {
                inQuotes = true;
            }
            else if (c == ',')
            {
                fields.Add(field.ToString());
                field.Clear();
            }
            else if (c == '\r')
            {
                if (reader.Peek() == '\n')
                    reader.Read();
                fields.Add(field.ToString());
                field.Clear();
                rows.Add(fields.ToArray());
                fields = new List<string>();
            }
            else if (c == '\n')
            {
                fields.Add(field.ToString());
                field.Clear();
                rows.Add(fields.ToArray());
                fields = new List<string>();
            }
            else
            {
                field.Append(c);
            }
        }

        return rows;
    }

    private static Encoding DetectEncoding(string path)
    {
        var bom = new byte[4];
        using (var fs = File.OpenRead(path))
        {
            _ = fs.Read(bom, 0, bom.Length);
        }

        if (bom[0] == 0xef && bom[1] == 0xbb && bom[2] == 0xbf)
            return Encoding.UTF8;
        if (bom[0] == 0xff && bom[1] == 0xfe)
            return Encoding.Unicode;
        if (bom[0] == 0xfe && bom[1] == 0xff)
            return Encoding.BigEndianUnicode;

        return Encoding.GetEncoding(1252);
    }
}
