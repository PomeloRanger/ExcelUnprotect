using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace ExcelUnprotect
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("No file was detected. Please drag and drop your Excel Files");
            }
            else
            {
                foreach (var files in args)
                {
                    if (CheckValidExtension(files))
                    {
                        RemoveProtection(files);
                    }
                }
            }
            Console.WriteLine("Please press any key to exit");
            Console.ReadKey();
        }

        private static void RemoveProtection(string files)
        {
            string tempPath = Path.GetTempPath() + "ExcelUnprotect";

            Directory.CreateDirectory(tempPath);
            File.Copy(files, tempPath + "\\excelworkbook.zip", true);

            ZipFile.ExtractToDirectory(tempPath + @"\excelworkbook.zip", tempPath + @"\excelworkbook", true);
            Console.WriteLine("Attempting to remove workbook protection");
            TryRemoveWorkbookProtection(tempPath + @"\excelworkbook\xl\workbook.xml");

            Console.WriteLine("Attempting to remove worksheet protection");
            string[] worksheetPath = Directory.GetFiles(tempPath + @"\excelworkbook\xl\worksheets");
            foreach (var worksheet in worksheetPath)
            {
                TryRemoveWorksheetProtection(worksheet);
            }
            ZipFile.CreateFromDirectory(tempPath + @"\excelworkbook", tempPath + @"\unprotectedworkbook.zip");

            //Save back to original file location with an added + _unprotected and change file extension to .xlsx
            Console.WriteLine("File is unprotected with the name " + files + "_unprotected.xlsx");
            File.Copy(tempPath + @"\unprotectedworkbook.zip", files + "_unprotected.xlsx", true);

            Console.WriteLine("Deleting Temp folder");
            Directory.Delete(tempPath, true);
        }


        private static bool CheckValidExtension(string file)
        {
            if (Path.GetExtension(file) == ".xlsx")
            {
                return true;
            }
            return false;
        }

        private static XmlDocument LoadDocument(string filePath)
        {
            XmlDocument document = new XmlDocument();
            document.Load(filePath);
            return document;
        }

        private static void TryRemoveWorksheetProtection(string filePath)
        {
            var doc = LoadDocument(filePath);
            RemoveWorksheetProtection(doc);
            doc.Save(filePath);
        }

        private static void TryRemoveWorkbookProtection(string filePath)
        {
            var doc = LoadDocument(filePath);
            RemoveWorkbookProtection(doc);
            doc.Save(filePath);
        }

        private static void RemoveWorksheetProtection(XmlDocument document)
        {
            foreach (XmlElement xmlElement in document.DocumentElement)
            {
                if (xmlElement.Name == "sheetProtection")
                {
                    xmlElement.ParentNode.RemoveChild(xmlElement);
                }
            }
        }

        private static void RemoveWorkbookProtection(XmlDocument document)
        {
            foreach (XmlNode items in document.DocumentElement.ChildNodes)
            {
                if (items.Name == "workbookProtection")
                {
                    items.ParentNode.RemoveChild(items);
                }
            }
        }
    }
}