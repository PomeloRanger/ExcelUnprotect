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
            LoadDocument(args);
        }

        private static void LoadDocument(string[] document)
        {
            if (document.Length == 0)
            {
                Console.WriteLine("No file was detected. Please drag and drop your Excel Files");
                Console.ReadLine();
            }
            else
            {
                RemoveProtection(document);
            }
        }

        private static void RemoveProtection(string[] directory)
        {
            foreach (var files in directory)
            {
                if (CheckValidExtension(files))
                {
                    Console.WriteLine("Unprotecting " + files);

                    Directory.CreateDirectory(Path.GetTempPath() + "ExcelUnprotect");

                    File.Copy(files, Path.GetTempPath() + "ExcelUnprotect\\excelworkbook.zip", true);

                    //Extract to folder so we can edit the
                    ZipFile.ExtractToDirectory(Path.GetTempPath() + "ExcelUnprotect\\excelworkbook.zip", Path.GetTempPath() + "ExcelUnprotect\\excelworkbook", true);

                    // In the directory /xl/workbook.xml if there is <workbookProtection> which will be there if workbook is under protection
                    TryRemoveWorkbookProtection(Path.GetTempPath() + "ExcelUnprotect\\excelworkbook\\xl\\workbook.xml");

                    //In the directory /xl/worksheets for every worksheet we check for <sheetProtection>
                    string[] worksheetPath = Directory.GetFiles(Path.GetTempPath() + "ExcelUnprotect\\excelworkbook\\xl\\worksheets");
                    foreach (var worksheet in worksheetPath)
                    {
                        TryRemoveWorksheetProtection(worksheet);
                    }

                    //Compile them back into zip
                    Console.WriteLine("Compiling into unprotected zip format");
                    ZipFile.CreateFromDirectory(Path.GetTempPath() + "ExcelUnprotect\\excelworkbook", Path.GetTempPath() + "ExcelUnprotect\\unprotectedworkbook.zip");

                    //Save back to original file location with an added + _unprotected and change file extension to .xlsx
                    Console.WriteLine("File is unprotected with the name " + files + "_unprotected.xlsx");
                    File.Copy(Path.GetTempPath() + "ExcelUnprotect\\unprotectedworkbook.zip", files + "_unprotected.xlsx", true);

                    //Delete folder from Temp path
                    Console.WriteLine("Deleting Temp folder");
                    Directory.Delete(Path.GetTempPath() + "ExcelUnprotect", true);
                }
                else
                {
                    Console.WriteLine("This file is not supported");
                }
            }
        }

        private static bool CheckValidExtension(string file)
        {
            if (Path.GetExtension(file) == ".xlsx")
            {
                return true;
            }
            return false;
        }

        private static void TryRemoveWorksheetProtection(string filePath)
        {
            XmlDocument document = new XmlDocument();
            document.Load(filePath);

            foreach (XmlElement xmlElement in document.DocumentElement)
            {
                if (xmlElement.Name == "sheetProtection")
                {
                    Console.WriteLine("Found sheetProtection at " + Path.GetFileName(filePath));
                    xmlElement.ParentNode.RemoveChild(xmlElement);
                }
            }
            document.Save(filePath);
        }

        private static void TryRemoveWorkbookProtection(string filePath)
        {
            XmlDocument document = new XmlDocument();
            document.Load(filePath);

            foreach (XmlNode items in document.DocumentElement.ChildNodes)
            {
                if (items.Name == "workbookProtection")
                {
                    Console.WriteLine("Found workbookProtection");
                    items.ParentNode.RemoveChild(items);
                }
            }
            document.Save(filePath);
        }
    }
}