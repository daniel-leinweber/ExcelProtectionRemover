using System.IO.Compression;
using System.Xml.Linq;

namespace ExcelProtectionRemover;

class Program
{
    static int Main(string[] args)
    {
        // Check for exactly two arguments: input and output file paths
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: ExcelProtectionRemover <inputFile.xlsx> <outputFile.xlsx>");

            return 1;
        }

        var inputFile = args[0];
        var outputFile = args[1];

        // Validate input file
        if (File.Exists(inputFile) == false)
        {
            Console.WriteLine($"Error: The file '{inputFile}' does not exist.");
            return 1;
        }

        if (Path.GetExtension(inputFile).ToLower() != ".xlsx")
        {
            Console.WriteLine("Error: Input file must have an .xlsx extension.");
            return 1;
        }

        // Create a temporary directory
        var tempDirectory = Path.Combine(
            Path.GetTempPath(),
            "ExcelProtection_" + Guid.NewGuid().ToString("N")
        );

        try
        {
            Directory.CreateDirectory(tempDirectory);

            // Extract the Excel (ZIP archive) file
            ZipFile.ExtractToDirectory(inputFile, tempDirectory);

            // Locate the worksheets directory
            var worksheetsDir = Path.Combine(tempDirectory, "xl", "worksheets");
            if (Directory.Exists(worksheetsDir) == false)
            {
                Console.WriteLine(
                    "Error: Could not find 'xl/worksheets' directory. Is this a valid .xlsx file?"
                );

                return 1;
            }

            // Process each worksheet XML file
            var xmlFiles = Directory.GetFiles(
                worksheetsDir,
                "*.xml",
                SearchOption.TopDirectoryOnly
            );

            foreach (var xmlFile in xmlFiles)
            {
                if (RemoveSheetProtection(xmlFile))
                {
                    Console.WriteLine($"Modified protection settings in: {xmlFile}");
                }
                else
                {
                    Console.WriteLine($"No protection tag found in: {xmlFile}");
                }
            }

            // Create the new .xlsx file (ZIP archive) from the modified directory
            if (File.Exists(outputFile))
            {
                File.Delete(outputFile);
            }

            ZipFile.CreateFromDirectory(tempDirectory, outputFile, CompressionLevel.Optimal, false);
            Console.WriteLine($"Processed workbook saved as: {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
            return 1;
        }
        finally
        {
            // Clean up the temporary directory
            try
            {
                Directory.Delete(tempDirectory, true);
            }
            catch
            { // Ignore clean-up errors
            }
        }

        return 0;
    }

    /// <summary>
    /// Removes all "sheetProtection" elements from the specified XML file.
    /// </summary>
    /// <param name="xmlFile">Path to the XML file.</param>
    /// <returns>True if one or more elements were removed; otherwise, false.</returns>
    static bool RemoveSheetProtection(string xmlFile)
    {
        var modified = false;

        try
        {
            var document = XDocument.Load(xmlFile);

            // Find all elements whose local name is "sheetProtection" (ignores namespaces)
            var protections = document
                .Descendants()
                .Where(x =>
                    x.Name.LocalName.Equals("sheetProtection", StringComparison.OrdinalIgnoreCase)
                )
                .ToList();

            if (protections.Any())
            {
                foreach (var protection in protections)
                {
                    protection.Remove();
                    modified = true;
                }

                // Save changes back to the XML file
                document.Save(xmlFile);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing '{xmlFile}': {ex.Message}");
        }

        return modified;
    }
}
