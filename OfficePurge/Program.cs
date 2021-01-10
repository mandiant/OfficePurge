using System;
using System.Collections.Generic;
using System.Linq;
using OpenMcdf;
using System.IO;
using System.IO.Compression;


namespace OfficePurge
{
	class Program
	{
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
						Console.WriteLine("\n[.] Will automatically decide which modules to purge.");
                    }
				}

                bool is_OpenXML = false;

                // Temp path to unzip OpenXML files to
                String unzipTempPath = "";

                string outFilename = Utils.getOutFilename(filename);
                string oleFilename = outFilename;

                // VBA Purging
                try
				{
					// Make a copy of document to VBA Purge if user is not listing  modules  
					if (!list_modules)
					{
						if (File.Exists(outFilename)) File.Delete(outFilename);
						File.Copy(filename, outFilename);
						filename = outFilename;
					}
					
                    try
                    {
                        unzipTempPath = CreateUniqueTempDirectory();
                        ZipFile.ExtractToDirectory(filename, unzipTempPath);

                        if (File.Exists(Path.Combine(unzipTempPath, "word", "vbaProject.bin"))) 
						{ 
							oleFilename = Path.Combine(unzipTempPath, "word", "vbaProject.bin"); 
						}
                        else if (File.Exists(Path.Combine(unzipTempPath, "xl", "vbaProject.bin"))) 
						{ 
							oleFilename = Path.Combine(unzipTempPath, "xl", "vbaProject.bin"); 
						}

                        is_OpenXML = true;
                    }
                    catch (Exception)
                    {
                        // Not OpenXML format, Maybe 97-2003 format, Make a copy
                        if (File.Exists(outFilename)) File.Delete(outFilename);
                        File.Copy(filename, outFilename);
					}

					CompoundFile cf = new CompoundFile(oleFilename, CFSUpdateMode.Update, 0);
					CFStorage commonStorage = cf.RootStorage;

					if (cf.RootStorage.TryGetStorage("Macros") != null)
					{
						commonStorage = cf.RootStorage.GetStorage("Macros");
					}
					
					if (cf.RootStorage.TryGetStorage("_VBA_PROJECT_CUR") != null)
					{
						commonStorage = cf.RootStorage.GetStorage("_VBA_PROJECT_CUR");
					}

					var vbaStorage = commonStorage.GetStorage("VBA");
					if(vbaStorage == null)
                    {
						throw new CFItemNotFound("Cannot find item");
					}

					// Grab data from "dir" module stream. Used to retrieve list of module streams in document.
					byte[] dirStream = Utils.Decompress(vbaStorage.GetStream("dir").GetData());
					List<Utils.ModuleInformation> vbaModules = Utils.ParseModulesFromDirStream(dirStream);

					// Only list module streams in document and return
					if (list_modules)
					{
						foreach (var vbaModule in vbaModules)
						{
							Console.WriteLine("[*] VBA module name: " + vbaModule.moduleName);
						}
						Console.WriteLine("[*] Finished listing modules\n");

                        return;
					}

					string [] dontPurgeTheseModules = {
						"ThisDocument", 
						"ThisWorkbook", 
						"Sheet",
					};

					byte[] streamBytes;
					bool module_found = false;
					foreach (var vbaModule in vbaModules)
					{
						//VBA Purging begins
						bool purge = true;

						if (module.Length > 0)
						{
							purge = vbaModule.moduleName == module;
						}
						else
						{
							foreach (string mod in dontPurgeTheseModules)
							{
								if (vbaModule.moduleName.StartsWith(mod)) purge = false;
							}
						}

						if (purge)
						{
							Console.WriteLine("\n[*] Purging VBA code in module: " + vbaModule.moduleName);
							Console.WriteLine("[*] Offset for code: " + vbaModule.textOffset);

							// Get the CompressedSourceCode from module   
							streamBytes = vbaStorage.GetStream(vbaModule.moduleName).GetData();
							string OG_VBACode = Utils.GetVBATextFromModuleStream(streamBytes, vbaModule.textOffset);

							// Remove P-code from module stream and set the module to only have the CompressedSourceCode
							streamBytes = Utils.RemovePcodeInModuleStream(streamBytes, vbaModule.textOffset, OG_VBACode);
							vbaStorage.GetStream(vbaModule.moduleName).SetData(streamBytes);
							module_found = true;
						}
					}

					if (module_found == false)
					{
						Console.WriteLine("\n[!] Could not find module in document (-m). List all module streams with (-l).\n");

						if (!is_OpenXML)
						{
							cf.Commit();
							cf.Close();
							CompoundFile.ShrinkCompoundFile(oleFilename);
							File.Delete(oleFilename);
							if (File.Exists(outFilename)) File.Delete(outFilename);
						}

                        return;
					}

					// Change offset to 0 so that document can find compressed source code.
					vbaStorage.GetStream("dir").SetData(Utils.Compress(Utils.ChangeOffset(dirStream)));
					Console.WriteLine("\n[*] Module offset changed to 0.");

					// Remove performance cache in _VBA_PROJECT stream. Replace the entire stream with _VBA_PROJECT header.

					string b1 = "00";
					string b2 = "00";

					Random rnd = new Random();
					b1 = String.Format("{0:X2}", rnd.Next(0, 255));
					b2 = String.Format("{0:X2}", rnd.Next(0, 255));

					byte[] data = Utils.HexToByte(String.Format("CC-61-FF-FF-00-{0}-{1}", b1, b2));
					vbaStorage.GetStream("_VBA_PROJECT").SetData(data);
					Console.WriteLine("[*] PerformanceCache removed from _VBA_PROJECT stream.");

					// Check if document contains SRPs. Must be removed for VBA Purging to work.
					try
					{
						for(int i = 0; i < 10; i++)
                        {
							string srp = String.Format("__SRP_{0}", i);
							var str = vbaStorage.TryGetStream(srp);
							if (str != null)
                            {
								vbaStorage.Delete(srp);
							}
						}
						
						Console.WriteLine("[*] SRP streams deleted!");
					}
					catch (Exception e)
					{
						Console.WriteLine("[*] No SRP streams found.");
					}

					// Commit changes and close
					cf.Commit();
					cf.Close();
					CompoundFile.ShrinkCompoundFile(oleFilename);

                    // Zip the file back up as a docm or xlsm
                    if (is_OpenXML)
                    {
						if (File.Exists(outFilename)) File.Delete(outFilename);
						ZipFile.CreateFromDirectory(unzipTempPath, outFilename);
                    }

                    Console.WriteLine("[+] VBA Purging completed successfully!\n");
				}

				// Error handle for file not found
				catch (FileNotFoundException ex) when (ex.Message.Contains("Could not find file"))
				{
					Console.WriteLine("[!] Could not find path or file (-f). \n");
				}

				// Error handle when document specified and file chosen don't match
				catch (CFItemNotFound ex) when (ex.Message.Contains("Cannot find item"))
				{
					Console.WriteLine("[!] File (-f) does not contain macros.\n");
				}

				// Error handle when document is not OLE/CFBF format
				catch (CFFileFormatException)
				{
					Console.WriteLine("[!] Incorrect filetype (-f). OfficePurge supports documents in .docm or .xlsm format as well as .doc/.xls/.pub in the Office 97-2003 format.\n");
				}
				finally
                {
                    if (is_OpenXML)
                    {
                        Directory.Delete(unzipTempPath, true);
                    }
                }
			}
			// Error handle for incorrect use of flags
			catch (IndexOutOfRangeException)
			{
				Console.WriteLine("\n[!] Flags (-d), (-f), (-m) need an argument. Make sure you have provided these flags an argument.\n");
			}
		}

        public static string CreateUniqueTempDirectory()
        {
            var uniqueTempDir = Path.GetFullPath(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString()));
            Directory.CreateDirectory(uniqueTempDir);
            return uniqueTempDir;
        }

    }
}
