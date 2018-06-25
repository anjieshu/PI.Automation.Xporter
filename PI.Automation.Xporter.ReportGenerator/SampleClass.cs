using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace PI.Automation.Xporter.ReportGenerator
{
	public class SampleClass
	{
		public SampleClass()
		{
		}

		private static List<String> filterColumnContent(String[] contents)
		{
			Console.WriteLine("Enter filterColumn()\n");
			if (contents == null)
			{
				Console.WriteLine("Input Error!\n");
				return null;
			}

			List<String> filteredColumn = new List<string>();
			foreach (String content in contents)
			{
				if (!filteredColumn.Contains(content))
				{
					filteredColumn.Add(content);
					Console.WriteLine("filtered value: {0}\n", content);
				}
			}

			Console.WriteLine("Exit filterColumn()\n");
			return filteredColumn;
		}

		private static String[] getColumnContent(ExcelWorksheet sheet, String columnName)
		{
			Console.WriteLine("Enter getColumnContent(): sheet={0}, column={1}\n", sheet, columnName);

			if (sheet == null || columnName == null)
			{
				Console.WriteLine("Input Error!");
				return null;
			}

			int columnIndex = getColumnIndex(sheet, columnName);
			String[] content = null;

			if (columnIndex > 0)
			{
				int rowLength = getSheetRowLength(sheet);
				if (rowLength > 1)
				{
					content = new String[rowLength - 1];
					for (int index = 0; index < rowLength - 1; index++)
					{
						content[index] = sheet.Cells[2 + index, columnIndex].Value.ToString();
						Console.WriteLine("get value: content[{0}]={1}\n", index, content[index]);
					}
				}
			}

			Console.WriteLine("Exit getColumnContent()\n");
			return content;
		}

		private static int getColumnIndex(ExcelWorksheet sheet, String columnName)
		{
			Console.WriteLine("Enter getColumnIndex(): sheet={0}, column={1}\n", sheet, columnName);

			if (sheet == null || columnName == null)
			{
				Console.WriteLine("Input Error!");
				return 0;
			}

			int columnLength = getSheetColumnLength(sheet);
			int columnIndex = 0;

			for (int index = 1; index <= columnLength; index++)
			{
				if (sheet.Cells[1, index].Value.ToString().Equals(columnName))
				{
					Console.WriteLine("{0} is found in column[{1}]!\n", columnName, index);
					columnIndex = index;
					break;
				}
			}

			if (columnIndex == 0)
			{
				Console.WriteLine("NO COLUMN FOUND !!!\n");
			}

			Console.WriteLine("Exit getColumnIndex()\n");
			return columnIndex;
		}

		private static ExcelWorksheet getSheet(ExcelPackage package, String sheetName)
		{
			Console.WriteLine("Enter getSheet(): package={0}, sheet={1}\n", package, sheetName);

			if (package == null || sheetName == null)
			{
				Console.WriteLine("Input Error!\n");
				return null;
			}

			ExcelWorksheets sheets = package.Workbook.Worksheets;
			ExcelWorksheet wantedSheet = null;
			foreach (ExcelWorksheet sheet in sheets)
			{
				if (sheet.Name.Equals(sheetName))
				{
					Console.WriteLine("{0} is found!\n", sheetName);
					wantedSheet = sheet;
				}
			}

			if (wantedSheet == null)
			{
				Console.WriteLine("NO SHEET FOUND !!!\n");
			}

			Console.WriteLine("Exit getSheet()\n");
			return wantedSheet;
		}

		public static int getSheetColumnLength(ExcelWorksheet sheet)
		{
			Console.WriteLine("Enter getSheetColumnLength(): sheet={0}\n", sheet.Name);

			int index = 1;
			while (index <= sheet.Dimension.End.Column)
			{
				//Console.WriteLine("Cells[1, {0}]={1}\n", index, sheet.Cells[1, index].Value);
				if (sheet.Cells[1, index].Value == null)
				{
					break;
				}

				index++;
			}

			Console.WriteLine("Exit getSheetColumnLength(): length={0}\n", index - 1);
			return index - 1;
		}

		public static int getSheetRowLength(ExcelWorksheet sheet)
		{
			Console.WriteLine("Enter getSheetRowLength(): sheet={0}\n", sheet.Name);

			int index = 1;
			while (index <= sheet.Dimension.End.Row)
			{
				//Console.WriteLine("Cells[{0}, 1]={1}\n", index, sheet.Cells[index, 1].Value);
				if (sheet.Cells[index, 1].Value == null)
				{
					break;
				}

				index++;
			}

			Console.WriteLine("Exit getSheetRowLength(): length={0}\n", index - 1);
			return index - 1;
		}

		public static void deleteColumnBetween(ExcelWorksheet sheet, String startColumn, String endColumn)
		{
			Console.WriteLine("Enter deleteColumnBetween({0}, {1}), sheet={2}\n", startColumn, endColumn, sheet.Name);

			int startIndex = getColumnIndex(sheet, startColumn);
			int endIndex = getColumnIndex(sheet, endColumn);

			if (startIndex > endIndex)
			{
				int tmp = startIndex;
				startIndex = endIndex;
				endIndex = tmp;
			}

			for (int index = startIndex + 1; index < endIndex; index++)
			{
				Console.WriteLine("Deleting column: {0} ...\n", sheet.Cells[1, index].Value);
				sheet.DeleteColumn(index);
			}

			Console.WriteLine("Exit deleteColumnBetween()\n");
		}

		private static void applySystltoSheet(ExcelWorksheet sheet)
		{
			var totalRow = sheet.Dimension.End.Row;
			var totalcolumn = sheet.Dimension.End.Column;
			sheet.Cells[1, 1, 3, totalcolumn].Style.Font.Bold = true;
			sheet.Cells[1, 1, 3, totalcolumn].Style.Fill.PatternType = ExcelFillStyle.Solid;
			sheet.Cells[1, 1, 3, totalcolumn].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.SkyBlue);
			sheet.Cells[1, 1, totalRow, totalcolumn].Style.Border.Top.Style = ExcelBorderStyle.Thin;
			sheet.Cells[1, 1, totalRow, totalcolumn].Style.Border.Left.Style = ExcelBorderStyle.Thin;
			sheet.Cells[1, 1, totalRow, totalcolumn].Style.Border.Right.Style = ExcelBorderStyle.Thin;
			sheet.Cells[1, 1, totalRow, totalcolumn].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
		}

		public static string Run(string fileName)
		{
			Console.WriteLine("Starting running...\nReading file: {0}", fileName);
			Console.WriteLine();

			FileInfo existingFile = new FileInfo(fileName);
			FileInfo newFile = new FileInfo(@"c:\temp\test_new.xlsx");
			using (ExcelPackage package = new ExcelPackage(existingFile))
			{
				// Get unique Revision values
				ExcelWorksheet testExecutionSheet = getSheet(package, "TestExecution");
				String[] contents = getColumnContent(testExecutionSheet, "Revision");
				List<String> columnRevision = filterColumnContent(contents);

				ExcelWorksheet testPlanSheet = getSheet(package, "TestPlan");

				// Delete columns between "Scenario" and "Defect link"
				deleteColumnBetween(testPlanSheet, "Scenario", "Defect link");

				// Insert columns between "Scenario" and "Defect link"
				var columnIndex = getColumnIndex(testPlanSheet, "Scenario");
				if (testPlanSheet != null && columnIndex > 0)
				{
					Console.WriteLine("Inserting columns...\n");
					testPlanSheet.InsertColumn(columnIndex + 1, columnRevision.Count);
					var totalRows = getSheetRowLength(testExecutionSheet);

					for (var index = 0; index < columnRevision.Count; index++)
					{
						testPlanSheet.Cells[1, columnIndex + 1 + index].Value = columnRevision[index];
					}

					Console.WriteLine("Columns inserted!\n");
				}

				Dictionary<String, DataModel> testStatusDic = new Dictionary<string, DataModel>();
				var count = getSheetRowLength(testExecutionSheet);
				var revisionColumnIndex = getColumnIndex(testExecutionSheet, "Revision");
				var testsColumnIndex = getColumnIndex(testExecutionSheet, "Tests");
				var finishedColumnIndex = getColumnIndex(testExecutionSheet, "Finished on");
				var testStatusColumnIndex = getColumnIndex(testExecutionSheet, "Test Status");

				for (var index = 2; index <= count; index++)
				{
					DataModel dataModel = new DataModel();
					Console.WriteLine("cells[{0},{1}]={2}", index, revisionColumnIndex,
						testExecutionSheet.Cells[index, revisionColumnIndex].Value);
					Console.WriteLine("cells[{0},{1}]={2}", index, testsColumnIndex,
						testExecutionSheet.Cells[index, testsColumnIndex].Value);
					Console.WriteLine("cells[{0},{1}]={2}", index, finishedColumnIndex,
						testExecutionSheet.Cells[index, finishedColumnIndex].Value);
					Console.WriteLine("cells[{0},{1}]={2}", index, testStatusColumnIndex,
						testExecutionSheet.Cells[index, testStatusColumnIndex].Value);

					dataModel.key = testExecutionSheet.Cells[index, revisionColumnIndex].Value +
					                testExecutionSheet.Cells[index, testsColumnIndex].Value.ToString();
					dataModel.timestamp = testExecutionSheet.Cells[index, finishedColumnIndex].Value.ToString();
					dataModel.status = testExecutionSheet.Cells[index, testStatusColumnIndex].Value.ToString();
					Console.WriteLine("Get dataModel: {0}", dataModel.toString());

					if (!testStatusDic.ContainsKey(dataModel.key))
					{
						Console.WriteLine("Key({0}) does NOT exist. Inserting ...", dataModel.key);
						testStatusDic[dataModel.key] = dataModel;
					}
					else
					{
						Console.WriteLine("Key({0}) already existed. Comparing the timestamp ...", dataModel.key);
						DataModel existData = testStatusDic[dataModel.key];
						if (dataModel.isFinishedLater(existData))
						{
							Console.WriteLine("dataModel({0}) is created later, replace the old one({1})", dataModel.toString(),
								existData.toString());
							testStatusDic[dataModel.key] = dataModel;
						}
					}
				}

				foreach (KeyValuePair<string, DataModel> entry in testStatusDic)
				{
					Console.WriteLine("{0}={1}-{2}", entry.Key, entry.Value.status, entry.Value.timestamp);
				}

				foreach (var t in columnRevision)
				{
					var revisionIndex = getColumnIndex(testPlanSheet, t);
					var testKeyIndex = getColumnIndex(testPlanSheet, "Test");
					if (testPlanSheet == null) continue;
					var rowCount = testPlanSheet.Dimension.End.Row;
					for (var x = 4; x <= rowCount; x++)
					{
						var testKey = testPlanSheet.Cells[1, revisionIndex].Value +
						              testPlanSheet.Cells[x, testKeyIndex].Value.ToString();
						testPlanSheet.Cells[x, revisionIndex].Value =
							testStatusDic.ContainsKey(testKey) ? testStatusDic[testKey].status : "n/a";
					}
				}

				applySystltoSheet(testPlanSheet);

				package.SaveAs(newFile);
			}

			Console.WriteLine("The End!\n");
			return newFile.Name;
		}
	}
}