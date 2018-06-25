using System;
using System.IO;

namespace PI.Automation.Xporter.ReportGenerator
{
	class Program
	{
		public static void Main(string[] args)
		{
			Console.WriteLine("Report generating...");
			try
			{
				//Set the output directory to the SampleApp folder where the app is running from. 
				Utils.OutputDir = new DirectoryInfo($"{AppDomain.CurrentDomain.BaseDirectory}SampleApp");
				Console.WriteLine("Running TestTool");
				var path = SampleClass.Run(@"c:\temp\test.xlsx");
				Console.WriteLine("The excel file is created: {0}", path);
				Console.WriteLine();
			}
			catch (Exception ex)
			{
				Console.WriteLine("Error: {0}", ex.Message);
			}

			var prevColor = Console.ForegroundColor;
			Console.ForegroundColor = ConsoleColor.Green;
			//Console.WriteLine($"Genereted sample workbooks can be found in {Utils.OutputDir.FullName}");
			Console.ForegroundColor = prevColor;

			Console.WriteLine();
			Console.WriteLine("Press the return key to exit...");
			Console.Read();
		}
	}
}