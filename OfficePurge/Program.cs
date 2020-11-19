using System;
using System.Collections.Generic;
using System.Linq;
using OpenMcdf;
using System.IO;


namespace OfficePurge
{
	class Program
	{
		private static string document = "";
		private static string filename = "";
		private static string module = "";
		private static bool list_modules = false;

		public static void PrintHelp()
		{
			Utils.HelpMenu();
		}
		static void Main(string[] args)
		{
			try
			{
				if (args.Length == 0 || args.Contains("-h"))
				{
					PrintHelp();
					return;
				}

				Dictionary<string, string> argDict = Utils.ParseArgs(args);

				if (argDict.ContainsKey("d"))
				{
					document = argDict["d"];
					if (document != "word" && document != "excel" && document != "publisher")
					{
						Console.WriteLine("\n[!] Unknown document type. Options are 'word', 'excel', or 'publisher'.\n");
						return;
					}
				}
				else
				{
					Console.WriteLine("\n[!] Missing document type (-d)\n");
					return;
				}

				if (argDict.ContainsKey("f"))
				{
					filename = argDict["f"];
				}
				else
				{
					Console.WriteLine("\n[!] Missing file (-f)\n");
					return;
				}

				if (args.Contains("-l"))
				{
					list_modules = true;
				}
				else
				{
					if (argDict.ContainsKey("m"))
					{
						module = argDict["m"];
					}
					else
					{
						Console.WriteLine("\n[!] Missing module (-m)\n");
						return;
					}
				}

				// VBA Purging
				try
				{
					// Make a copy of document to VBA Purge if user is not listing  modules  
					if (!list_modules)
					{
						string outFilename = Utils.getOutFilename(filename);
						string oleFilename = outFilename;

						if (File.Exists(outFilename)) File.Delete(outFilename);
						File.Copy(filename, outFilename);
						filename = outFilename;
					}

					CompoundFile cf = new CompoundFile(filename, CFSUpdateMode.Update, 0);
					CFStorage commonStorage = cf.RootStorage;

					if (document == "word")
					{
						commonStorage = cf.RootStorage.GetStorage("Macros");
					}

					else if (document == "excel")
					{
						commonStorage = cf.RootStorage.GetStorage("_VBA_PROJECT_CUR");
					}

					else if (document == "publisher")
					{
						commonStorage = cf.RootStorage.GetStorage("VBA");
					}


					// Grab data from "dir" module stream. Used to retrieve list of module streams in document.
					byte[] dirStream = Utils.Decompress(commonStorage.GetStorage("VBA").GetStream("dir").GetData());
					List<Utils.ModuleInformation> vbaModules = Utils.ParseModulesFromDirStream(dirStream);

					// Only list module streams in document and return
					if (list_modules)
					{
						foreach (var vbaModule in vbaModules)
						{
							Console.WriteLine("\n[*] VBA module name: " + vbaModule.moduleName);
						}
						Console.WriteLine("\n[*] Finished listing modules\n");
						return;
					}


					byte[] streamBytes;
					bool module_found = false;
					foreach (var vbaModule in vbaModules)
					{
						//VBA Purging begins
						if (vbaModule.moduleName == module)
						{
							Console.WriteLine("\n[*] VBA module name: " + vbaModule.moduleName);
							Console.WriteLine("\n[*] Offset for code: " + vbaModule.textOffset);
							Console.WriteLine("\n[*] Now purging VBA code in module: " + vbaModule.moduleName);

							// Get the CompressedSourceCode from module   
							streamBytes = commonStorage.GetStorage("VBA").GetStream(vbaModule.moduleName).GetData();
							string OG_VBACode = Utils.GetVBATextFromModuleStream(streamBytes, vbaModule.textOffset);

							// Remove P-code from module stream and set the module to only have the CompressedSourceCode
							streamBytes = Utils.RemovePcodeInModuleStream(streamBytes, vbaModule.textOffset, OG_VBACode);
							commonStorage.GetStorage("VBA").GetStream(vbaModule.moduleName).SetData(streamBytes);
							module_found = true;

						}
					}

					if (module_found == false)
					{
						Console.WriteLine("\n[!] Could not find module in document (-m). List all module streams with (-l).\n");
						cf.Commit();
						cf.Close();
						CompoundFile.ShrinkCompoundFile(filename);
						File.Delete(filename);
						return;
					}


					// Change offset to 0 so that document can find compressed source code.
					commonStorage.GetStorage("VBA").GetStream("dir").SetData(Utils.Compress(Utils.ChangeOffset(dirStream)));
					Console.WriteLine("\n[*] Module offset changed to 0.");

					// Remove performance cache in _VBA_PROJECT stream. Replace the entire stream with _VBA_PROJECT header.
					byte[] data = Utils.HexToByte("CC-61-FF-FF-00-00-00");
					commonStorage.GetStorage("VBA").GetStream("_VBA_PROJECT").SetData(data);
					Console.WriteLine("\n[*] PerformanceCache removed from _VBA_PROJECT stream.");

					// Check if document contains SRPs. Must be removed for VBA Purging to work.
					try
					{
						commonStorage.GetStorage("VBA").Delete("__SRP_0");
						commonStorage.GetStorage("VBA").Delete("__SRP_1");
						commonStorage.GetStorage("VBA").Delete("__SRP_2");
						commonStorage.GetStorage("VBA").Delete("__SRP_3");
						Console.WriteLine("\n[*] SRP streams deleted!");
					}
					catch (Exception)
					{
						Console.WriteLine("\n[*] No SRP streams found.");
					}

					// Commit changes and close
					cf.Commit();
					cf.Close();
					CompoundFile.ShrinkCompoundFile(filename);
					Console.WriteLine("\n[*] VBA Purging completed successfully!\n");
				}

				// Error handle for file not found
				catch (FileNotFoundException ex) when (ex.Message.Contains("Could not find file"))
				{
					Console.WriteLine("\n[!] Could not find path or file (-f). \n");
				}

				// Error handle when document specified and file chosen don't match
				catch (CFItemNotFound ex) when (ex.Message.Contains("Cannot find item"))
				{
					Console.WriteLine("\n[!] File (-f) does not match document type selected (-d).\n");
				}

				// Error handle when document is not OLE/CFBF format
				catch (CFFileFormatException)
				{
					Console.WriteLine("\n[!] Incorrect filetype (-f). Must be an OLE strucutred file. OfficePurge supports .doc, .xls, or .pub documents.\n");
				}
			}
			// Error handle for incorrect use of flags
			catch (IndexOutOfRangeException)
			{
				Console.WriteLine("\n[!] Flags (-d), (-f), (-m) need an argument. Make sure you have provided these flags an argument.\n");
			}
		}
	}
}
