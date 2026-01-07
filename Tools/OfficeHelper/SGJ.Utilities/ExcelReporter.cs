using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SGJ.Utilities
{
    public class ExcelReporter
    {
        #region Worksheet
        public static WorksheetPart CopySheet(SpreadsheetDocument pobjDoc, string pstrSourceSheetName, string pstrClonedSheetName)
        {
            WorkbookPart workbookPart = pobjDoc.WorkbookPart;
            //Get the source sheet to be copied
            WorksheetPart sourceSheetPart = GetWorkSheetPartByName(workbookPart, pstrSourceSheetName);
            if (sourceSheetPart == null)
            {
                sourceSheetPart = GetFirstOrDefaultWorkSheetPart(workbookPart);
            }
            //------------------------------------------------------
            SpreadsheetDocument tempDoc = SpreadsheetDocument.Create(new MemoryStream(), pobjDoc.DocumentType);
            WorkbookPart tempWorkbookPart = tempDoc.AddWorkbookPart();
            WorksheetPart tempWorksheetPart = tempWorkbookPart.AddPart<WorksheetPart>(sourceSheetPart);
            //------------------------------------------------------
            //Add cloned sheet and all associated parts to workbook
            WorksheetPart clonedWorksheetPart = workbookPart.AddPart<WorksheetPart>(tempWorksheetPart);
            //Table definition parts are somewhat special and need unique ids...so let's make an id based on count
            int numTableDefParts = sourceSheetPart.GetPartsCountOfType<TableDefinitionPart>();
            int tableId = numTableDefParts;
            //Clean up table definition parts (tables need unique ids)
            if (numTableDefParts != 0)
            {
                FixupTableParts(clonedWorksheetPart, numTableDefParts);
            }
            //There should only be one sheet that has focus
            CleanView(clonedWorksheetPart);
            //Add new sheet to main workbook part
            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            Sheet copiedSheet = new Sheet();
            copiedSheet.Name = pstrClonedSheetName;
            copiedSheet.Id = workbookPart.GetIdOfPart(clonedWorksheetPart);
            copiedSheet.SheetId = (uint)sheets.ChildElements.Count + 1;
            sheets.Append(copiedSheet);
            return clonedWorksheetPart;
        }

        public static void RemoveSheet(SpreadsheetDocument pobjDoc, string pstrSheetName)
        {
            WorkbookPart workbookPart = pobjDoc.WorkbookPart;
            //Get the source sheet to be copied
            var worksheetPart = GetWorkSheetPartByName(workbookPart, pstrSheetName);
            var sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name.Value.Equals(pstrSheetName)).First();
            if (sheet == null)
            {
                return;
            }
            else
            {
                sheet.Remove();
                workbookPart.DeletePart(worksheetPart);
            }
        }

        public static WorksheetPart GetFirstOrDefaultWorkSheetPart(WorkbookPart pobjWorkbookPart)
        {
            //Get the relationship id of the sheetname
            string relId = pobjWorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault().Id;
            return (WorksheetPart)pobjWorkbookPart.GetPartById(relId);
        }

        public static WorksheetPart GetWorkSheetPartByName(WorkbookPart pobjWorkbookPart, string pstrSheetName)
        {
            //Get the relationship id of the sheetname
            string relId = pobjWorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name.Value.Equals(pstrSheetName)).First().Id;
            return (WorksheetPart)pobjWorkbookPart.GetPartById(relId);
        }

        public static WorksheetPart GetWorkSheetPartById(WorkbookPart pobjWorkbookPart, string pstrId)
        {
            return (WorksheetPart)pobjWorkbookPart.GetPartById(pstrId);
        }

        public static void CleanView(WorksheetPart pobjWorksheetPart)
        {
            //There can only be one sheet that has focus
            SheetViews views = pobjWorksheetPart.Worksheet.GetFirstChild<SheetViews>();
            if (views != null)
            {
                views.Remove();
                pobjWorksheetPart.Worksheet.Save();
            }
        }

        public static void FixupTableParts(WorksheetPart pobjWorksheetPart, int pintTableDefParts)
        {
            //Every table needs a unique id and name
            foreach (TableDefinitionPart tableDefPart in pobjWorksheetPart.TableDefinitionParts)
            {
                pintTableDefParts++;
                tableDefPart.Table.Id = (uint)pintTableDefParts;
                tableDefPart.Table.DisplayName = "CopiedTable" + pintTableDefParts;
                tableDefPart.Table.Name = "CopiedTable" + pintTableDefParts;
                tableDefPart.Table.Save();
            }
        }

        #endregion

        #region Row
        public static int GetRowIndex(string cellReference)
        {
            var regex = new Regex("[0-9]+");
            var match = regex.Match(cellReference);
            return Int32.Parse(match.Value);
        }

        public static Row CreateOrGetRow(Worksheet pobjWorksheet, uint rowIndex)
        {
            Row row = pobjWorksheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            if (row == null)
            {
                row = new Row { RowIndex = rowIndex };
            }
            return row;
        }

        #endregion

        #region Column

        public static string GetColumnName(string cellReference)
        {
            var regex = new Regex("[A-Za-z]+");
            var match = regex.Match(cellReference);
            return match.Value;
        }

        public static string ConvertColumnNumberToName(int intCol)
        {
            var intFirstLetter = ((intCol) / 676) + 64;
            var intSecondLetter = ((intCol % 676) / 26) + 64;
            var intThirdLetter = (intCol % 26) + 65;

            var firstLetter = (intFirstLetter > 64) ? (char)intFirstLetter : ' ';
            var secondLetter = (intSecondLetter > 64) ? (char)intSecondLetter : ' ';
            var thirdLetter = (char)intThirdLetter;

            return string.Concat(firstLetter, secondLetter, thirdLetter).Trim();
        }

        public static string ConvertColumnNumberToName1(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public static int ConvertColumnNameToNumber(string columnName)
        {
            var alpha = new Regex("^[A-Z]+$");
            if (!alpha.IsMatch(columnName)) throw new ArgumentException();

            char[] colLetters = columnName.ToCharArray();
            Array.Reverse(colLetters);

            var convertedValue = 0;
            for (int i = 0; i < colLetters.Length; i++)
            {
                char letter = colLetters[i];
                int current = i == 0 ? letter - 65 : letter - 64; // ASCII 'A' = 65
                convertedValue += current * (int)Math.Pow(26, i);
            }
            return convertedValue;
        }

        #endregion

        #region Cell

        public static Cell GetCell(Row row, string columnName)
        {
            if (row != null)
            {
                return row.Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + row.RowIndex, true) == 0).First();
            }
            else
            {
                return null;
            }
        }

        public static Cell GetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            Row row = CreateOrGetRow(worksheet, rowIndex);
            if (row != null)
            {
                return row.Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0).First();
            }
            else
            {
                return null;
            }
        }

        public static Cell CreateOrUpdateCell(Row row, string columnName, string text)
        {
            Cell cell = GetCell(row, columnName);
            if (cell != null)
            {
                cell.CellValue = new CellValue(text);
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
            }
            else
            {
                cell = new Cell { DataType = CellValues.InlineString, CellReference = columnName + row.RowIndex };
                var inlineString = new InlineString();
                var t = new Text { Text = text };
                inlineString.AppendChild(t);
                cell.AppendChild(inlineString);
            }
            return cell;
        }

        public static Cell CreateOrUpdateCell(WorksheetPart pobjWorksheetPart, string columnName, uint rowIndex, string text)
        {
            if (pobjWorksheetPart != null)
            {
                Cell cell = GetCell(pobjWorksheetPart.Worksheet, columnName, rowIndex);
                if (cell != null)
                {
                    cell.CellValue = new CellValue(text);
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                }
                else
                {
                    cell = new Cell { DataType = CellValues.InlineString, CellReference = columnName + rowIndex };
                    var inlineString = new InlineString();
                    var t = new Text { Text = text };
                    inlineString.AppendChild(t);
                    cell.AppendChild(inlineString);
                }
                return cell;
            }
            return null;
        }


        public static void MergeCells(WorksheetPart worksheetPart, Cell cell, int RowSpan)
        {
            Worksheet workSheet = worksheetPart.Worksheet;
            MergeCells mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().First();
            if (mergeCells == null)
            {
                mergeCells = new MergeCells();
                workSheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());
            }

            //mergeCells.Append(new MergeCell() { Reference = new StringValue("C1:F1") });
            string columnName = GetColumnName(cell.CellReference);
            int columnIndex = ConvertColumnNameToNumber(columnName);
            int rowIndex = GetRowIndex(cell.CellReference);
            string value = string.Format("{1}{0}:{2}{0}", rowIndex, ConvertColumnNumberToName(columnIndex), ConvertColumnNumberToName(columnIndex + RowSpan - 1));
            mergeCells.Append(new MergeCell() { Reference = new StringValue(value) });

        }

        #endregion

    }
}
