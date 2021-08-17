using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Excel2DTDL
{
    public class ExcelParser
    {
        DTDLModel dtdlModel;
        List<Interface> interfaces;

        public ExcelParser()
        {
            dtdlModel = new DTDLModel();
            interfaces = new List<Interface>();
        }

        public DTDLModel Parse(string filename)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filename)))
            {
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    if (worksheet.Name.ToLower() == "sample") continue;

                    var newInterface = new Interface()
                    {
                        context = "dtmi:dtdl:context;2",
                        type = "Interface"
                    };

                    int rowCount = worksheet.Dimension.End.Row;     //get row count
                    // Get interface
                    int rowNum = 1;
                    if (worksheet.Cells[rowNum, 1].Value == null || worksheet.Cells[rowNum, 1].Value.ToString() != "interface")
                    {
                        Log.Error("Cell A1 must be a interface.");
                    }
                    rowNum = 2;
                    newInterface.id = worksheet.Cells[rowNum, 2].Value.ToString();
                    newInterface.name = worksheet.Cells[rowNum, 3].Value == null ? null : worksheet.Cells[rowNum, 3].Value.ToString();
                    newInterface.displayName = worksheet.Cells[rowNum, 4].Value == null ? null : worksheet.Cells[rowNum, 4].Value.ToString();
                    newInterface.extends = worksheet.Cells[rowNum, 5].Value == null ? null : worksheet.Cells[rowNum, 5].Value.ToString();
                    newInterface.description = worksheet.Cells[rowNum, 6].Value == null ? null : worksheet.Cells[rowNum, 6].Value.ToString();

                    rowNum = 3;
                    string currentContent = "";
                    var contents = new List<Content>();
                    do
                    {
                        string contentType;
                        
                        if (worksheet.Cells[rowNum, 1].Value == null)
                        {
                            contentType = currentContent;
                            switch (contentType.ToLower())
                            {
                                case "property":
                                    contents.Add(GenerateProperty(worksheet, rowNum));
                                    break;
                                case "telemetry":
                                    contents.Add(GenerateTelemetry(worksheet, rowNum));
                                    break;
                                case "component":
                                    contents.Add(GenerateComponent(worksheet, rowNum));
                                    break;
                                case "relationship":
                                    contents.Add(GenerateRelationship(worksheet, rowNum));
                                    break;
                            }
                            rowNum += 1;

                            continue;
                        }
                        else
                        {
                            currentContent = worksheet.Cells[rowNum, 1].Value.ToString().Trim();
                            rowNum += 1;
                            continue;
                        }
                    } while (rowNum <= rowCount);
                    newInterface.contents = contents.ToArray();
                    interfaces.Add(newInterface);
                }

                dtdlModel.InterfaceArray = interfaces.ToArray();
                return dtdlModel;
            }
        }

        private Property GenerateProperty(ExcelWorksheet sheet, int rowNum)
        {
            var prop = new Property()
            {
                name = sheet.Cells[rowNum, 2].Value.ToString(),
                schema = sheet.Cells[rowNum, 3].Value.ToString(),
                writable = sheet.Cells[rowNum, 4].Value == null ? (bool?)null : bool.Parse(sheet.Cells[rowNum, 4].Value.ToString().ToLower()),
                type = sheet.Cells[rowNum, 5].Value == null ? new string[1] { "Property"} : new string[2] { "Property", sheet.Cells[rowNum, 5].Value.ToString() },
                unit = sheet.Cells[rowNum, 6].Value == null ? null : sheet.Cells[rowNum, 6].Value.ToString()
            };

            return prop;
        }

        private Telemetry GenerateTelemetry(ExcelWorksheet sheet, int rowNum)
        {
            var telemetry = new Telemetry()
            {
                name = sheet.Cells[rowNum, 2].Value.ToString(),
                schema = sheet.Cells[rowNum, 3].Value.ToString(),
                type = sheet.Cells[rowNum, 4].Value == null ? new string[1] { "Telemetry" } : new string[2] { "Property", sheet.Cells[rowNum, 4].Value.ToString() },
                unit = sheet.Cells[rowNum, 5].Value == null ? null : sheet.Cells[rowNum,5].Value.ToString()
            };

            return telemetry;
        }

        private Component GenerateComponent(ExcelWorksheet sheet, int rowNum)
        {
            var comp = new Component()
            {
                name = sheet.Cells[rowNum, 2].Value.ToString(),
                schema = sheet.Cells[rowNum, 3].Value.ToString(),
                type =  new string[1] { "Component" }, 
                displayName = sheet.Cells[rowNum, 4].Value == null ? null : sheet.Cells[rowNum, 4].Value.ToString(),
                description = sheet.Cells[rowNum, 5].Value == null ? null : sheet.Cells[rowNum, 5].Value.ToString(),
            };

            return comp;
        }

        private RelationShip GenerateRelationship(ExcelWorksheet sheet, int rowNum)
        {
            var relation = new RelationShip()
            {
                name = sheet.Cells[rowNum, 2].Value.ToString(),
                type = new string[1] { "Relationship" },
                displayName = sheet.Cells[rowNum, 3].Value == null ? null : sheet.Cells[rowNum, 3].Value.ToString(),
                target = sheet.Cells[rowNum, 4].Value.ToString(),
            };

            return relation;
        }
    }
}
