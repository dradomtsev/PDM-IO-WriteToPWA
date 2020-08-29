using System;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Specialized;
using System.Text.RegularExpressions;
using PDM_IO_WriteToPWA_ReadXLSX.Model;

using Newtonsoft.Json;

using Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client;

using System.Security;
using System.Net;

using PWAWebLogin;
using Microsoft.Online.SharePoint.MigrationCenter.Common;

using System.Globalization;

namespace PDM.IO.PWA.TaskECF
{
    class Program
    {
        private const string SiteURL = "https://archimatika.sharepoint.com/sites/pwatest";
        private static ProjectContext projContext = new ProjectContext(SiteURL);
        private static readonly CultureInfo invCult = CultureInfo.InvariantCulture;
        static void Main()
        {
            string fileName = @"C:\PRJ-Tasks-CustomFields.xlsx";
            string sheetName = "Tasks";
            string dataTableName = "Tasks";
            string ProjectFieldName = "ProjectId";
            string TaskFieldName = "TaskId";
            string customFieldNameSource = "Draft_TaskCountEstimatedCost";
            string customFieldNameTarget = "TaskCountEstimatedCost";

            DataforOutput dataExport = new DataforOutput
            {
                data = new List<dataExportItem>()
            };

            DataforOutput dataImport = new DataforOutput
            {
                data = new List<dataExportItem>()
            };

            string DataJSON = "";

            ReadExcelTableData(ref dataExport, fileName, sheetName, dataTableName, ProjectFieldName, TaskFieldName, customFieldNameSource);
            WriteToJson(ref dataExport,ref DataJSON);
            ReadFromJson(ref dataImport, ref DataJSON);
            WorkPWA(ref dataImport, ref customFieldNameTarget);
        }

        static void WorkPWA(ref DataforOutput dataImport, ref string customFieldNameTarget)
        {
            // Connect to Sharepoint using cookies
            var cookies = WebLogin.GetWebLoginCookie(new Uri(SiteURL));
            projContext.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e)
            {
                e.WebRequestExecutor.WebRequest.CookieContainer = new CookieContainer();
                e.WebRequestExecutor.WebRequest.CookieContainer.SetCookies(new Uri(SiteURL), cookies);
            };

            // Write data to PWA
            WriteToPWATaskECF(ref dataImport, customFieldNameTarget);
        }

        private static void WriteToPWATaskECF(ref DataforOutput dataImport, string customFieldNameTarget)
        {
            // Get Project ID. TODO: Group records by project to write data by project groups
            string sprojGUID = dataImport.data.FirstOrDefault().ProjectID; // "0ee2dab3-b6bd-e911-8159-3085a9af61fc";
            Guid projGUID = Guid.Parse(sprojGUID);

            // Get project by ProjectId
            var projects = projContext.LoadQuery(projContext.Projects.Where(p => p.Id == projGUID).Include(p => p.Id, p => p.Name));
            projContext.Load(projContext.CustomFields);
            projContext.ExecuteQuery();
            var project = projects.FirstOrDefault();

            // Get custom field internal name
            var cfInternalName = projContext.CustomFields.FirstOrDefault(q => q.Name == customFieldNameTarget).InternalName;

            // Checkout project for modification
            var draftProject = project.CheckOut();
            projContext.Load(draftProject, p => p.Tasks.Include(t => t.Id, t => t.Name));
            projContext.ExecuteQuery();

            // Get all tasks for project
            var tasks = draftProject.Tasks;

            string sTaskGUID = "";
            foreach (var record in dataImport.data)
            {
                // Get Task ID from data import
                sTaskGUID = record.TaskID; // "43e2dab3-b6bd-e911-8159-3085a9af61fc";
                Guid taskGUID = Guid.Parse(sTaskGUID);

                // Get Task for project by TaskId
                var draftTask = tasks.FirstOrDefault(q => q.Id == taskGUID);

                // Main result: write data to Task ECF
                draftTask[cfInternalName] = record.TaskCountEstimatedCost;
            }

            // Publish project
            draftProject.Publish(true);
            projContext.ExecuteQuery();
        }

        static void ReadFromJson(ref DataforOutput dataImport, ref string DataJSON)
        {
            dataImport = JsonConvert.DeserializeObject<DataforOutput>(DataJSON);
        }

        static void WriteToJson(ref DataforOutput dataExport, ref string DataJSON)
        {
            DataJSON = JsonConvert.SerializeObject(dataExport);
        }

        static void ReadExcelTableData(ref DataforOutput dataExport, string fileName, string sheetName, string dataTableName, string ProjectFieldName, string TaskFieldName, string customFieldName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                // Get basic data strucure from Excel
                WorkbookPart workbookPart = document.WorkbookPart;
                Workbook workBook = workbookPart.Workbook;
                Sheet sheet = workBook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
                SharedStringTable sharedStringTable = workBook.WorkbookPart.SharedStringTablePart.SharedStringTable;
                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);
                Worksheet workSheet = worksheetPart.Worksheet;
                TableDefinitionPart tableDefinitionPart = worksheetPart.TableDefinitionParts.FirstOrDefault(r => r.Table.Name == dataTableName);
                Table excelTable = tableDefinitionPart.Table;

                // Get table range
                var cellRange = excelTable.Reference;
                uint startCell = GetRowIndex(cellRange.Value.Split(':')[0]);
                uint endCell = GetRowIndex(cellRange.Value.Split(':')[1]);

                

                // Set Cell names
                Cell ProjectIDHeaderCell = workSheet.Descendants<Row>().Where(r => r.RowIndex.Value == startCell).FirstOrDefault().Elements<Cell>().Where(c => String.Compare(GetCellValuebyID(ref sharedStringTable, short.Parse(c.CellValue.Text, invCult)), ProjectFieldName, true, invCult) == 0).FirstOrDefault();
                string ProjectIDHeaderName = GetColumnName(ProjectIDHeaderCell.CellReference.Value);

                Cell TaskIDHeaderCell = workSheet.Descendants<Row>().Where(r => r.RowIndex.Value == startCell).FirstOrDefault().Elements<Cell>().Where(c => String.Compare(GetCellValuebyID(ref sharedStringTable, short.Parse(c.CellValue.Text, invCult)), TaskFieldName, true, invCult) == 0).FirstOrDefault();
                string TaskIDHeaderName = GetColumnName(TaskIDHeaderCell.CellReference.Value);

                Cell customFieldHeaderCell = workSheet.Descendants<Row>().Where(r => r.RowIndex.Value == startCell).FirstOrDefault().Elements<Cell>().Where(c => String.Compare(GetCellValuebyID(ref sharedStringTable, short.Parse(c.CellValue.Text, invCult)), customFieldName, true, invCult) == 0).FirstOrDefault();
                string CustomFieldHeaderName = GetColumnName(customFieldHeaderCell.CellReference.Value);
                
                // Iterate through filtered rows in table
                foreach (Row row in workSheet.Descendants<Row>().Where(r => r.RowIndex.Value > startCell && r.RowIndex.Value <= endCell && (r.Hidden == null || r.Hidden.Value != true)))
                {
                    dataExportItem itemDataExport = new dataExportItem();

                    foreach (Cell cell in row)
                    {
                        string columnName = GetColumnName(cell.CellReference.Value);

                        // Write data to internal POCO object
                        if (columnName == ProjectIDHeaderName)
                            itemDataExport.ProjectID = GetCellValuebyID(ref sharedStringTable, short.Parse(cell.CellValue.Text, invCult));
                        if (columnName == TaskIDHeaderName)
                            itemDataExport.TaskID = GetCellValuebyID(ref sharedStringTable, short.Parse(cell.CellValue.Text, invCult));
                        if (columnName == CustomFieldHeaderName)
                            itemDataExport.TaskCountEstimatedCost = float.Parse(cell.CellValue.Text, invCult);
                    }
                    // Add data to POCO object
                    dataExport.data.Add(itemDataExport);
                }
            }
        }

        // Given a cell name, parses the specified cell to get the row index.
        private static uint GetRowIndex(string cellName)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            return uint.Parse(match.Value, invCult);
        }

        // Given a cell name, parses the specified cell to get the column name.
        private static string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }

        // Get cell value by ID from Shared Strings table
        private static string GetCellValuebyID(ref SharedStringTable sharedStringTable, int cellID)
        {
            return sharedStringTable.Elements<SharedStringItem>().ElementAt(cellID).Text.InnerText;
        }
    }
}
