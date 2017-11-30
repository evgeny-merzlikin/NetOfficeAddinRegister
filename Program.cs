/*
 * Created by SharpDevelop.
 * User: merzlev
 * Date: 29.11.2017
 * Time: 16:12
 * 
 */
using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.IO;
using Microsoft.Build.Utilities;

namespace OfficeAddinRegister
{
	class Program
	{
		private static string FullAddinName;
		
		
		public static void Main(string[] args)
		{
			//Console.WriteLine("Current dir: {0}",System.Environment.CurrentDirectory);
			if (args.Length < 1) {
				PrintHelpAndExit();
			}
			
			string AddinName = args[0];
			string RegFile = String.Format("{0}.reg", AddinName);
			string OfficeApp="";
			
			if (args.Length == 1)
			{
				string ProgName = System.Diagnostics.Process.GetCurrentProcess().ProcessName;
				// try to get office app name from exe name
				if (ProgName.ToLower().Contains("word")) {
					OfficeApp = "Word";
					
				} else if (ProgName.ToLower().Contains("excel")) {
					OfficeApp = "Excel";
					
				} else if (ProgName.ToLower().Contains("outlook")) {
					OfficeApp = "Outlook";
					
				} else if (ProgName.ToLower().Contains("access")) {
					OfficeApp = "  Access";
					
				} else if (ProgName.ToLower().Contains("powerpoint")) {
					OfficeApp = "PowerPoint";
					
				} else if (ProgName.ToLower().Contains("msproject")) {
					OfficeApp = "  MSProject";

				} else if (ProgName.ToLower().Contains("    visio")) {
					OfficeApp = "  Visio";
					
				} else {
					PrintHelpAndExit();
				}


				
			} else {
			  OfficeApp = args[1];
		  }
			
			
			Console.WriteLine("Addin: {0}\r\nOfficeApp: {1}\r\n", AddinName, OfficeApp);
			
			if (!File.Exists(AddinName)) {
				Console.Error.WriteLine("{0} file does not exist", AddinName);
				Console.ReadKey(true);
				Environment.Exit(1);
 			}
			
			FullAddinName = Path.GetFullPath(AddinName);
			
			Console.WriteLine("Searching regasm.exe...");
			string regasm = Path.Combine(ToolLocationHelper.GetPathToDotNetFramework(TargetDotNetFrameworkVersion.VersionLatest), "regasm.exe");
			
			if (!File.Exists(regasm)) {
				Console.Error.WriteLine("Critial Error: cannot locate regasm.exe");
				Console.ReadKey(true);
				Environment.Exit(1);
 			}
			
			Console.WriteLine("Found {0}\r\n", regasm);
			
			string command = string.Format("\"{0}\" /nologo /codebase /regfile:{1}", AddinName, RegFile);
			Console.WriteLine("Executing regasm.exe {0} ...", command);

	    Process process = new Process();
	    process.StartInfo.FileName = regasm;
	    process.StartInfo.Arguments = command; // Note the /c command (*)
	    process.StartInfo.UseShellExecute = false;
	    process.StartInfo.RedirectStandardOutput = true;
	    process.StartInfo.RedirectStandardError = true;
	    process.Start();
	    //* Read the output (or the error)
	    string output = process.StandardOutput.ReadToEnd();
	    //Console.WriteLine(output);
	    string err = process.StandardError.ReadToEnd();
	    
	    process.WaitForExit();
			int result = process.ExitCode;
			if (result != 0) {

				Console.Error.WriteLine(err);
				Console.ReadKey(true);
				Environment.Exit(1);
				
			}
			
			if (!File.Exists(RegFile)) {
				Console.Error.WriteLine("{0} file does not exist", RegFile);
				Console.ReadKey(true);
				Environment.Exit(1);
 			}
			
			
      string regdata = File.ReadAllText(RegFile);
			
			Console.WriteLine(output);
			
			//Console.WriteLine(regdata);
			
			//Console.WriteLine("regasm.exe completed successfully!");
			
			string[] chunks = regdata.Split(new string[] {"\r\n\r\n"}, StringSplitOptions.RemoveEmptyEntries);
			
			StringBuilder sb = new StringBuilder();
			sb.AppendLine("Windows Registry Editor Version 5.00\r\n");
      
      string ProgID = Regex.Match(chunks[1],@"\[HKEY_CLASSES_ROOT\\(?<ProgID>[^\]\\]+)\]").Groups["ProgID"].Value; 

      sb.AppendFormat("[HKEY_CURRENT_USER\\Software\\Microsoft\\Office\\{0}\\Addins\\{1}]\r\n", OfficeApp, ProgID);
	    sb.AppendFormat("\"Description\"=\"{0}\"\r\n", ProgID);
	    sb.AppendFormat("\"FriendlyName\"=\"{0}\"\r\n", ProgID);
			sb.AppendFormat("\"LoadBehavior\"=dword:00000003\r\n\r\n");


      sb.AppendFormat("[HKEY_CURRENT_USER\\Software\\Wow6432Node\\Microsoft\\Office\\{0}\\Addins\\{1}]\r\n", OfficeApp, ProgID);
	    sb.AppendFormat("\"Description\"=\"{0}\"\r\n", ProgID);
	    sb.AppendFormat("\"FriendlyName\"=\"{0}\"\r\n", ProgID);
			sb.AppendFormat("\"LoadBehavior\"=dword:00000003\r\n\r\n");



			for (int i = 1 ; i < chunks.Length; i++) {
				//Console.WriteLine("{0}: {1}", i, chunks[i]);
				if (Regex.IsMatch(
					chunks[i], 
					@"^\[HKEY_CLASSES_ROOT\\CLSID\\")) 
				{
					
					sb.AppendLine(ReplaceCodebase(
						Regex.Replace(
							chunks[i],
							@"^\[HKEY_CLASSES_ROOT\\CLSID\\",			
							@"[HKEY_CURRENT_USER\Software\Classes\CLSID\")));
					
					sb.AppendLine();
					
					sb.AppendLine(ReplaceCodebase(
						Regex.Replace(
							chunks[i],
							@"^\[HKEY_CLASSES_ROOT\\CLSID\\",	
							@"[HKEY_CURRENT_USER\Software\Classes\Wow6432Node\CLSID\")));
 
					sb.AppendLine();
					
				} else if (Regex.IsMatch(
					chunks[i], 
					@"^\[HKEY_CLASSES_ROOT\\")) 
				{
				
					sb.AppendLine(ReplaceCodebase(
						Regex.Replace(
							chunks[i],
							@"^\[HKEY_CLASSES_ROOT\\",					
							@"[HKEY_CURRENT_USER\Software\Classes\")));

					sb.AppendLine();
				} else {
					sb.AppendLine(chunks[i]);
					sb.AppendLine();
				}
			}
			string InstallRegFile = String.Format("Install {0}", RegFile);
			Console.WriteLine("The registry information has been added to file {0}.", InstallRegFile);		
			//Console.WriteLine(sb);		
			File.WriteAllText(InstallRegFile, sb.ToString());
			File.Delete(RegFile);
			Console.ReadKey(true);
		}
		

		public static string ReplaceCodebase(string text)
		{
		
			if (Regex.IsMatch(text, @"\""CodeBase\""=.*"))
			{
			  string CodeBase = Regex.Replace(
					FullAddinName, @"\\", @"\\");
				return Regex.Replace(
					text,
					@"\""CodeBase\""=.*",					
					String.Format(@"""CodeBase""=""{0}""", CodeBase));
			} else {
				return text;
			}

			
		}
		
		public static void PrintHelpAndExit()
		{
			Console.WriteLine("This program generates .reg file for per-user registration of office addin (NetOffice library). ");
	
			Console.Error.WriteLine("Too few arguments. Usage: OfficeAddinRegister.exe <addin.dll> (Word|Excel|PowerPoint|Outlook...etc)");
	
			Console.Error.WriteLine("You can omit the second parameter and put it in the name of this exe: rename OfficeAddinRegister to WordAddinRegister.exe, for example.");
	
			Console.ReadKey(true);
			Environment.Exit(1);

				
		}
		
	}
}
