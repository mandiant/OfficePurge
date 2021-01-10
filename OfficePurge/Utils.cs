using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Kavod.Vba.Compression;

namespace OfficePurge
{
    class Utils
    {
		public static Dictionary<string, string> ParseArgs(string[] args)
		{
			Dictionary<string, string> ret = new Dictionary<string, string>();
			if (args.Length > 0)
			{
				for (int i = 0; i < args.Length; i += 2)
				{
					if (args[i].Substring(1).ToLower() == "l")
					{
						ret.Add(args[i].Substring(1).ToLower(), "true");
					}
					else
					{
						ret.Add(args[i].Substring(1).ToLower(), args[i + 1]);
					}
				}
			}
			return ret;
		}
		public static void HelpMenu()
		{
			Console.WriteLine("\n  __  ____  ____  __  ___  ____  ____  _  _  ____   ___  ____ ");
			Console.WriteLine(" /  \\(  __)(  __)(  )/ __)(  __)(  _ \\/ )( \\(  _ \\ / __)(  __)");
			Console.WriteLine("(  O )) _)  ) _)  )(( (__  ) _)  ) __/) \\/ ( )   /( (_ \\ ) _) ");
			Console.WriteLine(" \\__/(__)  (__)  (__)\\___)(____)(__)  \\____/(__\\_) \\___/(____) v1.0");
			Console.WriteLine("\n\n Author: Andrew Oliveau\n");
			Console.WriteLine(" DESCRIPTION:");
			Console.WriteLine("\n\tOfficePurge is a C# tool that VBA purges malicious Office documents. ");
			Console.WriteLine("\tVBA purging removes P-code from module streams within Office documents. ");
			Console.WriteLine("\tDocuments that only contain source code and no compiled code are more");
			Console.WriteLine("\tlikely to evade AV detection and YARA rules.\n\n");
			Console.WriteLine(" USAGE:");
			Console.WriteLine("\t-f : Filename to VBA Purge");
			Console.WriteLine("\t-m : Module within document to VBA Purge");
			Console.WriteLine("\t-l : List module streams in document");
			Console.WriteLine("\t-h : Show help menu.\n");
			Console.WriteLine(" EXAMPLES:");
			Console.WriteLine("\n\t .\\OfficePurge.exe -d word -f .\\malicious.doc -m NewMacros");
			Console.WriteLine("\t .\\OfficePurge.exe -d excel -f .\\payroll.xls -m Module1");
			Console.WriteLine("\t .\\OfficePurge.exe -d publisher -f .\\donuts.pub -m ThisDocument");
			Console.WriteLine("\t .\\OfficePurge.exe -d word -f .\\malicious.doc -l\n");
		}
		public static List<ModuleInformation> ParseModulesFromDirStream(byte[] dirStream)
		{
			// 2.3.4.2 dir Stream: Version Independent Project Information
			// https://msdn.microsoft.com/en-us/library/dd906362(v=office.12).aspx
			// Dir stream is ALWAYS in little endian

			List<ModuleInformation> modules = new List<ModuleInformation>();

			int offset = 0;
			UInt16 tag;
			UInt32 wLength;
			ModuleInformation currentModule = new ModuleInformation { moduleName = "", textOffset = 0 };

			while (offset < dirStream.Length)
			{
				tag = GetWord(dirStream, offset);
				wLength = GetDoubleWord(dirStream, offset + 2);

				// taken from Pcodedmp
				if (tag == 9)
					wLength = 6;
				else if (tag == 3)
					wLength = 2;

				switch (tag)
				{
					// MODULESTREAMNAME Record
					case 26:
						currentModule.moduleName = System.Text.Encoding.UTF8.GetString(dirStream, (int)offset + 6, (int)wLength);
						break;

					// MODULEOFFSET Record
					case 49:
						currentModule.textOffset = GetDoubleWord(dirStream, offset + 6);
						modules.Add(currentModule);
						currentModule = new ModuleInformation { moduleName = "", textOffset = 0 };
						break;
				}

				offset += 6;
				offset += (int)wLength;
			}

			return modules;
		}

		public class ModuleInformation
		{
			// Name of VBA module stream
			public string moduleName;

			// Offset of VBA CompressedSourceCode in VBA module stream
			public UInt32 textOffset;
		}

		public static UInt16 GetWord(byte[] buffer, int offset)
		{
			var rawBytes = new byte[2];
			Array.Copy(buffer, offset, rawBytes, 0, 2);
			return BitConverter.ToUInt16(rawBytes, 0);
		}

		public static UInt32 GetDoubleWord(byte[] buffer, int offset)
		{
			var rawBytes = new byte[4];
			Array.Copy(buffer, offset, rawBytes, 0, 4);
			return BitConverter.ToUInt32(rawBytes, 0);
		}
		public static byte[] Compress(byte[] data)
		{
			var buffer = new DecompressedBuffer(data);
			var container = new CompressedContainer(buffer);
			return container.SerializeData();
		}
		public static byte[] Decompress(byte[] data)
		{
			var container = new CompressedContainer(data);
			var buffer = new DecompressedBuffer(container);
			return buffer.Data;
		}
		public static string GetVBATextFromModuleStream(byte[] moduleStream, UInt32 textOffset)
		{
			string vbaModuleText = Encoding.UTF8.GetString(Decompress(moduleStream.Skip((int)textOffset).ToArray()));
			return vbaModuleText;
		}
		public static byte[] RemovePcodeInModuleStream(byte[] moduleStream, UInt32 textOffset, string OG_VBACode)
		{
			return Compress(Encoding.UTF8.GetBytes(OG_VBACode)).ToArray();
		}
		public static string getOutFilename(String filename)
		{
			string fn = Path.GetFileNameWithoutExtension(filename);
			string ext = Path.GetExtension(filename);
			string path = Path.GetDirectoryName(filename);
			return Path.Combine(path, fn + "_PURGED" + ext);
		}
		public static byte[] HexToByte(string hex)
		{
			hex = hex.Replace("-", "");
			byte[] raw = new byte[hex.Length / 2];
			for (int i = 0; i < raw.Length; i++)
			{
				raw[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
			}
			return raw;
		}
		public static byte[] ChangeOffset(byte[] dirStream)
		{
			int offset = 0;
			UInt16 tag;
			UInt32 wLength;

			// Change MODULEOFFSET to 0
			string zeros = "\0\0\0\0";
			
			while (offset < dirStream.Length)
			{
				tag = GetWord(dirStream, offset);
				wLength = GetDoubleWord(dirStream, offset + 2);

				// taken from Pcodedmp
				if (tag == 9)
					wLength = 6;
				else if (tag == 3)
					wLength = 2;

				switch (tag)
				{
					// MODULEOFFSET Record
					case 49:
						uint offset_change = GetDoubleWord(dirStream, offset + 6);
						UTF8Encoding encoding = new UTF8Encoding();
						encoding.GetBytes(zeros, 0, (int)wLength, dirStream, (int)offset + 6);
						break;
				}

				offset += 6;
				offset += (int)wLength;
			}
			return dirStream;
		}
	}
}
