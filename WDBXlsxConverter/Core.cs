using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;
using System.Threading;

namespace WDBXlsxConverter
{
    internal class Core
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("");

            if (args.Length == 0)
            {
                WDBMethods.ErrorExit("Enough arguments not specified. please launch the program with -? switch for more info.");
            }

            if (args[0] == "-?")
            {
                Console.WriteLine("Game Codes:");
                Console.WriteLine("-ff131 = Sets conversion compatibility to FFXIII");
                Console.WriteLine("-ff132 = Sets conversion compatibility to FFXIII-2 and LR");

                Console.WriteLine("");
                Console.WriteLine("Tool actions:");
                Console.WriteLine("WDBXlsxConverter -? = Display this help page");
                Console.WriteLine("WDBXlsxConverter -gamecode \"wdbFilePath\" = Converts the wdb data into a new xlsx file");
                Console.WriteLine("WDBXlsxConverter -gamecode \"wdbFilePath\" -i = Converts the wdb data into a new xlsx file without the fieldnames (only when gamecode is -ff131)");

                Console.WriteLine("");
                Console.WriteLine("Examples:");
                Console.WriteLine("WDBXlsxConverter.exe -?");
                Console.WriteLine("WDBXlsxConverter.exe -ff131 \"auto_clip.wdb\"");
                Console.WriteLine("WDBXlsxConverter.exe -ff131 \"auto_clip.wdb\" -i");
                Console.WriteLine("WDBXlsxConverter.exe -ff132 \"auto_clip.wdb\"");

                Console.ReadLine();
                Environment.Exit(0);
            }

            if (args.Length < 2)
            {
                WDBMethods.ErrorExit("Enough arguments not specified for this process");
            }


            var gameCode = args[0];
            var inFile = args[1];

            try
            {
                if (!File.Exists(inFile))
                {
                    WDBMethods.ErrorExit("Specified file in the argument is missing");
                }

                using (var wdbReader = new BinaryReader(File.Open(inFile, FileMode.Open, FileAccess.Read)))
                {
                    _ = wdbReader.BaseStream.Position = 0;
                    if (wdbReader.ReadBytesString(3, false) != "WPD")
                    {
                        WDBMethods.ErrorExit("Not a valid WPD file");
                    }

                    _ = wdbReader.BaseStream.Position += 1;

                    switch (gameCode)
                    {
                        case "-ff131":
                            var wdbVarsXIII = new XIII.WDBVariablesXIII
                            {
                                IgnoreKnown = args.Length > 2 && args.Contains("-i"),
                                RecordCount = wdbReader.ReadBytesUInt32(true)
                            };

                            if (wdbVarsXIII.RecordCount == 0)
                            {
                                WDBMethods.ErrorExit("No records/sections are present in this file");
                            }

                            XIII.SectionsParser.MainSections(wdbReader, wdbVarsXIII);

                            Console.WriteLine("");
                            Console.WriteLine($"Total records: {wdbVarsXIII.RecordCount}");
                            Console.WriteLine("");

                            Console.WriteLine("Parsing records....");
                            Console.WriteLine("");
                            Thread.Sleep(1000);

                            wdbVarsXIII.WDBName = Path.GetFileNameWithoutExtension(inFile);
                            wdbVarsXIII.xlsxName = Path.Combine(Path.GetDirectoryName(inFile), wdbVarsXIII.WDBName + ".xlsx");

                            using (var workbook = new XLWorkbook())
                            {
                                XIII.SectionsParser.MainSectionsToWB(wdbVarsXIII, workbook);
                                XIII.RecordsParser.ProcessRecords(wdbReader, wdbVarsXIII, workbook);
                            }
                            break;

                        case "-ff132":
                            var wdbVarsXIII2LR = new XIII2LR.WDBVariablesXIII2LR();

                            wdbVarsXIII2LR.WDBName = Path.GetFileNameWithoutExtension(inFile);
                            wdbVarsXIII2LR.xlsxName = Path.Combine(Path.GetDirectoryName(inFile), wdbVarsXIII2LR.WDBName + ".xlsx");

                            wdbVarsXIII2LR.RecordCount = wdbReader.ReadBytesUInt32(true);

                            if (wdbVarsXIII2LR.RecordCount == 0)
                            {
                                WDBMethods.ErrorExit("No records/sections are present in this file");
                            }

                            XIII2LR.SectionsParser.MainSections(wdbReader, wdbVarsXIII2LR);

                            Console.WriteLine("");
                            Console.WriteLine($"Total records: {wdbVarsXIII2LR.RecordCount}");
                            Console.WriteLine("");

                            Console.WriteLine("Parsing records....");
                            Console.WriteLine("");
                            Thread.Sleep(1000);

                            using (var workbook = new XLWorkbook())
                            {
                                XIII2LR.SectionsParser.MainSectionsToWB(wdbVarsXIII2LR, workbook);
                                XIII2LR.RecordsParser.ProcessRecords(wdbReader, wdbVarsXIII2LR, workbook);
                            }
                            break;

                        default:
                            WDBMethods.ErrorExit("Specified gamecode is invalid");
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("");
                Console.WriteLine("An exception has occured!");
                Console.WriteLine("");
                Console.WriteLine($"{ex}");

                Environment.Exit(2);
            }

            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("Finished converting records to xlsx file");
            Console.ReadLine();

            Environment.Exit(0);
        }
    }
}