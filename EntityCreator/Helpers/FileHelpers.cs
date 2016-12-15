using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using EntityCreator.Models;

namespace EntityCreator.Helpers
{
    public static class FileHelpers
    {
        private const string lineSeperator = "##########################################################################################################################";
        public static void WriteErrorLog(Dictionary<string, List<Exception>> errors)
        {
            foreach (var error in errors)
            {
                WriteLog(error.Value, "_Error",  error.Key);
            }
        }

        public static void WriteWarningLog(Dictionary<string, List<Exception>> warnings)
        {
            foreach (var warning in warnings)
            {
                WriteLog(warning.Value, "_Warning", warning.Key);
            }
        }

        private static void WriteLog(IEnumerable<Exception> issueList, string postfix, string fileNameFull)
        {
            var fileNameSplitted = fileNameFull.Split('.');
            var fullFileName = fileNameFull.Insert(fileNameFull.IndexOf("." + fileNameSplitted.Last(), StringComparison.Ordinal), postfix);

            //var fileNameSplitted = fileNameFull.Split('\\');
            //var fileExtensionSplitted = fileNameSplitted.Last().Split('.');
            //var fileName = fileNameSplitted.Last().Insert("." + fileExtensionSplitted.Last(), "")
            //var fullFileName = fileNameFull.Replace(fileNameSplitted.Last(), fileName);
            var errorMessage = string.Empty;
            foreach (var exception in issueList)
            {
                errorMessage += exception + "\n" + lineSeperator + "\n";
            }

            var path = fullFileName + ".txt";
            using (var sw = new StreamWriter(path, true))
            {
                sw.WriteLine(errorMessage);
                sw.Close();
            }
        }

        public static IEnumerable<string> GetExcelFiles(string selectedPath)
        {
            return Directory.GetFiles(selectedPath, "*.xlsx");
        }

        public static EntityTemplate GetEntityTemplateFromFile(string excelFile)
        {
            var excelHelper = new ExcelHelper(excelFile);
            return excelHelper.GetEntityTemplateFromFile(excelFile);
        }

        public static void ExtractResources()
        {
            var local = Environment.CurrentDirectory;
            ExtractEmbeddedResource(local, "EntityCreator.Resources", "Template.xlsx");
            ExtractEmbeddedResource(local, "EntityCreator.Resources", "HowToUse.docx");
        }

        private static void ExtractEmbeddedResource(string outputDir, string resourceLocation, string file)
        {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceLocation + @"." + file))
            {
                if (stream != null)
                {
                    using (var fileStream = new FileStream(Path.Combine(outputDir, file), FileMode.Create))
                    {
                        for (var i = 0; i < stream.Length; i++)
                        {
                            fileStream.WriteByte((byte)stream.ReadByte());
                        }

                        fileStream.Close();
                    }
                }
            }
        }

        public static void MarkFileAsProcessed(string excelFile)
        {
            var file = excelFile.Split('.');
            File.Move(excelFile, file[0] + "_processed." + file[1]);
        }
    }
}