using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SGJ.Utilities;

namespace Test
{
    class TestExcel
    {

        static void Test(string[] args)
        {
            List<EDRIDesignDocument> lcolDrawings = new List<EDRIDesignDocument>() { 
            new EDRIDesignDocument("ddd", " ffffffffffffff","Arch", "00"),
            new EDRIDesignDocument("xx"," ffffffffffffff", "BIM", "00"),
            new EDRIDesignDocument("qq"," ffffffffffffff", "BIM", "00"),
            new EDRIDesignDocument("zz"," ffffffffffffff", "Arch", "00"),
            new EDRIDesignDocument("dsddd"," ffffffffffffff", "Arch", "00"),
            new EDRIDesignDocument("ddsdasdd"," ffffffffffffff", "VVV", "00"),
            new EDRIDesignDocument("ddsdccd"," ffffffffffffff", "VVV", "00")
                };
            ProcessExcelReport(lcolDrawings, "1.xlsx");
        }

        public static void ProcessExcelReport(List<EDRIDesignDocument> lcolDocument, string pstrFile)
        {
            try
            {
                EDRIDrawingList lobjDrawingList = PrepareData(lcolDocument, "t", "q");
                //open file
                var lobjDoc = SpreadsheetDocument.Open(pstrFile, true);
                var workbookPart = lobjDoc.WorkbookPart;
                Dictionary<string, WorksheetPart> lcolWorksheetPart = new Dictionary<string, WorksheetPart>();
                // copy sheet
                foreach (var lobjDocument in lcolDocument)
                {
                    //worksheetPart
                    if (!lcolWorksheetPart.ContainsKey(lobjDocument.DisciplineName))
                    {
                        WorksheetPart worksheetPart = ExcelReporter.CopySheet(lobjDoc, "Sample", lobjDocument.DisciplineName);
                        lcolWorksheetPart.Add(lobjDocument.DisciplineName, worksheetPart);
                    }
                }
                // write discipline
                foreach (var lobjSheet in lcolWorksheetPart)
                {
                    //
                    string disciplineName = lobjSheet.Key;
                    EDRIDiscipline discipline = lobjDrawingList.Disciplines[disciplineName];

                    var worksheetPart = lobjSheet.Value;
                    var workSheet = worksheetPart.Worksheet;
                    var sheetData = workSheet.Elements<SheetData>().First();
                    var rows = sheetData.Elements<Row>().ToList();
                    var columns = workSheet.Descendants<Columns>().FirstOrDefault();
                    // write discipline header
                    ExcelReporter.CreateOrUpdateCell(worksheetPart, "C", 1, lobjDrawingList.ProjectName);
                    ExcelReporter.CreateOrUpdateCell(worksheetPart, "C", 2, lobjDrawingList.ProjectNo);
                    ExcelReporter.CreateOrUpdateCell(worksheetPart, "C", 3, disciplineName);
                    //write buildings 
                    uint lintBodyIndex = 6;
                    foreach (var lobjBuilding in discipline.Buildings)
                    {
                        //write building header
                        Row row = ExcelReporter.CreateOrGetRow(workSheet, lintBodyIndex);
                        Cell cell = ExcelReporter.CreateOrUpdateCell(row, "B", lobjBuilding.Key);
                        ExcelReporter.MergeCells(worksheetPart, cell, 2);
                        lintBodyIndex += 1;
                        // write document detail
                        int lintBuildingIndex = 1;
                        foreach (var lobjDocument in lobjBuilding.Value.Documents)
                        {
                            Row inrow = ExcelReporter.CreateOrGetRow(workSheet, lintBodyIndex);
                            ExcelReporter.CreateOrUpdateCell(inrow, "A", lintBuildingIndex.ToString());
                            ExcelReporter.CreateOrUpdateCell(inrow, "B", lobjDocument.DrawingNo);
                            ExcelReporter.CreateOrUpdateCell(inrow, "C", lobjDocument.Description);
                            lintBuildingIndex += 1;
                            lintBodyIndex += 1;
                        }
                    }

                    // Save the worksheet.
                    worksheetPart.Worksheet.Save();
                }
                // remove sample sheet
                ExcelReporter.RemoveSheet(lobjDoc, "Sample");
                // save file
                workbookPart.Workbook.Save();
                lobjDoc.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static EDRIDrawingList PrepareData(List<EDRIDesignDocument> pcolDocument, string pstrProjectName, string pstrProjectNo)
        {
            EDRIDrawingList lobjDrawingList = new EDRIDrawingList();
            foreach (var lobjDocument in pcolDocument)
            {
                if (string.IsNullOrEmpty(lobjDocument.DisciplineName) || string.IsNullOrEmpty(lobjDocument.BuildingName))
                {
                    continue;
                }
                //discipline
                if (!lobjDrawingList.Disciplines.ContainsKey(lobjDocument.DisciplineName))
                {
                    EDRIDiscipline discipline = new EDRIDiscipline();
                    discipline.Name = lobjDocument.DisciplineName;
                    discipline.Description = lobjDocument.DisciplineName;
                    lobjDrawingList.Disciplines.Add(discipline.Name, discipline);
                }
                //building
                if (!lobjDrawingList.Disciplines[lobjDocument.DisciplineName].Buildings.ContainsKey(lobjDocument.BuildingName))
                {
                    EDRIBuilding building = new EDRIBuilding();
                    building.Name = lobjDocument.BuildingName;
                    building.Description = lobjDocument.BuildingName;
                    lobjDrawingList.Disciplines[lobjDocument.DisciplineName].Buildings.Add(building.Name, building);
                }
                //document
                lobjDrawingList.Disciplines[lobjDocument.DisciplineName].Buildings[lobjDocument.BuildingName].Documents.Add(lobjDocument);
            }
            return lobjDrawingList;
        }

    }
}
