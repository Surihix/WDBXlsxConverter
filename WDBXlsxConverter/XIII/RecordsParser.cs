using ClosedXML.Excel;
using System;
using System.IO;

namespace WDBXlsxConverter.XIII
{
    internal class RecordsParser
    {
        public static void ProcessRecords(BinaryReader wdbReader, WDBVariablesXIII wdbVars, XLWorkbook workbook)
        {
            var cellX = 1;  // vertical
            var cellY = 1;  // horizontal

            // Write all the fields in the 
            // worksheet
            IXLWorksheet mainSheet;

            if (wdbVars.IsKnown)
            {
                mainSheet = workbook.Worksheets.Add(wdbVars.SheetName);
                SharedMethods.WriteToSheet(mainSheet, cellX, cellY, "Record", 2, true);
                cellY++;

                foreach (var fieldCell in wdbVars.Fields)
                {
                    SharedMethods.WriteToSheet(mainSheet, cellX, cellY, fieldCell, 2, true);
                    cellY++;
                }

                cellX = 2;

                ParseRecordsWithFields(wdbReader, wdbVars, mainSheet, cellX);
            }
            else
            {
                mainSheet = workbook.Worksheets.Add(wdbVars.WDBName);
                SharedMethods.WriteToSheet(mainSheet, cellX, cellY, "Record", 2, true);
                cellY++;

                ParseRecordsWithoutFields(wdbVars, mainSheet, cellX, cellY, wdbReader);
            }

            // Save the workbook as
            // excel file
            if (File.Exists(wdbVars.xlsxName))
            {
                File.Delete(wdbVars.xlsxName);
            }

            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("Writing new worksheet data to excel file....");

            mainSheet.Rows().AdjustToContents();
            mainSheet.Columns().AdjustToContents();

            workbook.SaveAs(wdbVars.xlsxName);
        }


        private static void ParseRecordsWithFields(BinaryReader wdbReader, WDBVariablesXIII wdbVars, IXLWorksheet mainSheet, int cellX)
        {
            // Process each record's data
            var sectionPos = wdbReader.BaseStream.Position;
            string currentRecordName;
            byte[] currentRecordData;
            var strTypeListIndex = 0;
            var currentRecordDataIndex = 0;
            int cellY;

            for (int r = 0; r < wdbVars.RecordCount; r++)
            {
                _ = wdbReader.BaseStream.Position = sectionPos;
                currentRecordName = wdbReader.ReadBytesString(16, false);

                cellY = 1;
                Console.WriteLine($"Record: {currentRecordName}");
                SharedMethods.WriteToSheet(mainSheet, cellX, cellY, currentRecordName, 2, false);
                cellY++;

                currentRecordData = SharedMethods.SaveSectionData(wdbReader, false);

                for (int f = 0; f < wdbVars.FieldCount; f++)
                {
                    switch (wdbVars.StrtypelistValues[strTypeListIndex])
                    {
                        // bitpacked
                        case 0:
                            var binaryData = BitOperationHelpers.UIntToBinary(SharedMethods.DeriveUIntFromSectionData(currentRecordData, currentRecordDataIndex, true));
                            var binaryDataIndex = binaryData.Length;
                            var fieldBitsToProcess = 32;

                            int iTypedataVal;
                            uint uTypeDataVal;
                            int fTypeDataVal;

                            while (fieldBitsToProcess != 0 && f < wdbVars.FieldCount)
                            {
                                var fieldType = wdbVars.Fields[f].Substring(0, 1);
                                var fieldNum = SharedMethods.DeriveFieldNumber(wdbVars.Fields[f]);

                                switch (fieldType)
                                {
                                    // sint
                                    case "i":
                                        if (fieldNum == 0)
                                        {
                                            iTypedataVal = BitOperationHelpers.BinaryToInt(binaryData, binaryDataIndex - 32, 32);
                                            fieldBitsToProcess = 0;

                                            Console.WriteLine($"{wdbVars.Fields[f]}: {iTypedataVal}");
                                            SharedMethods.WriteToSheet(mainSheet, cellX, cellY, iTypedataVal, 4, false);
                                            cellY++;

                                            break;
                                        }
                                        if (fieldNum > fieldBitsToProcess)
                                        {
                                            f--;
                                            fieldBitsToProcess = 0;
                                            continue;
                                        }
                                        else
                                        {
                                            binaryDataIndex -= fieldNum;

                                            iTypedataVal = BitOperationHelpers.BinaryToInt(binaryData, binaryDataIndex, fieldNum);
                                            fieldBitsToProcess -= fieldNum;

                                            Console.WriteLine($"{wdbVars.Fields[f]}: {iTypedataVal}");
                                            SharedMethods.WriteToSheet(mainSheet, cellX, cellY, iTypedataVal, 4, false);
                                            cellY++;

                                            if (fieldBitsToProcess != 0)
                                            {
                                                f++;
                                            }
                                        }
                                        break;

                                    // uint 
                                    case "u":
                                        if (fieldNum == 0)
                                        {
                                            uTypeDataVal = BitOperationHelpers.BinaryToUInt(binaryData, binaryDataIndex - 32, 32);
                                            fieldBitsToProcess = 0;

                                            Console.WriteLine($"{wdbVars.Fields[f]}: {uTypeDataVal}");
                                            SharedMethods.WriteToSheet(mainSheet, cellX, cellY, uTypeDataVal, 3, false);
                                            cellY++;

                                            break;
                                        }
                                        if (fieldNum > fieldBitsToProcess)
                                        {
                                            f--;
                                            fieldBitsToProcess = 0;
                                            continue;
                                        }
                                        else
                                        {
                                            binaryDataIndex -= fieldNum;

                                            uTypeDataVal = BitOperationHelpers.BinaryToUInt(binaryData, binaryDataIndex, fieldNum);
                                            fieldBitsToProcess -= fieldNum;

                                            Console.WriteLine($"{wdbVars.Fields[f]}: {uTypeDataVal}");
                                            SharedMethods.WriteToSheet(mainSheet, cellX, cellY, uTypeDataVal, 3, false);
                                            cellY++;

                                            if (fieldBitsToProcess != 0)
                                            {
                                                f++;
                                            }
                                        }
                                        break;

                                    // float (bitpacked as int) 
                                    case "f":
                                        if (fieldNum == 0)
                                        {
                                            fTypeDataVal = BitOperationHelpers.BinaryToInt(binaryData, binaryDataIndex - 32, 32);
                                            fieldBitsToProcess = 0;

                                            Console.WriteLine($"{wdbVars.Fields[f]}: {fTypeDataVal}");
                                            SharedMethods.WriteToSheet(mainSheet, cellX, cellY, fTypeDataVal, 4, false);
                                            cellY++;

                                            break;
                                        }
                                        if (fieldNum > fieldBitsToProcess)
                                        {
                                            f--;
                                            fieldBitsToProcess = 0;
                                            continue;
                                        }
                                        else
                                        {
                                            binaryDataIndex -= fieldNum;

                                            fTypeDataVal = BitOperationHelpers.BinaryToInt(binaryData, binaryDataIndex, fieldNum);
                                            fieldBitsToProcess -= fieldNum;

                                            Console.WriteLine($"{wdbVars.Fields[f]}: {fTypeDataVal}");
                                            SharedMethods.WriteToSheet(mainSheet, cellX, cellY, fTypeDataVal, 4, false);
                                            cellY++;

                                            if (fieldBitsToProcess != 0)
                                            {
                                                f++;
                                            }
                                        }
                                        break;
                                }
                            }

                            strTypeListIndex++;
                            currentRecordDataIndex += 4;
                            break;

                        // float value
                        case 1:
                            var floatDataVal = SharedMethods.DeriveFloatFromSectionData(currentRecordData, currentRecordDataIndex, true);

                            Console.WriteLine($"{wdbVars.Fields[f]}: {floatDataVal}");
                            SharedMethods.WriteToSheet(mainSheet, cellX, cellY, floatDataVal, 1, false);
                            cellY++;

                            strTypeListIndex++;
                            currentRecordDataIndex += 4;
                            break;

                        // !!string section offset
                        case 2:
                            var stringDataOffset = SharedMethods.DeriveUIntFromSectionData(currentRecordData, currentRecordDataIndex, true);
                            var derivedString = SharedMethods.DeriveStringFromArray(wdbVars.StringsData, (int)stringDataOffset);

                            if (derivedString == "")
                            {
                                derivedString = "{null}";
                            }

                            Console.WriteLine($"{wdbVars.Fields[f]}: {derivedString}");
                            SharedMethods.WriteToSheet(mainSheet, cellX, cellY, derivedString, 2, false);
                            cellY++;

                            strTypeListIndex++;
                            currentRecordDataIndex += 4;
                            break;

                        // uint value
                        case 3:
                            if (wdbVars.Fields[f].StartsWith("u64"))
                            {
                                var processArray = new byte[8];
                                Array.ConstrainedCopy(currentRecordData, currentRecordDataIndex, processArray, 0, 8);
                                Array.Reverse(processArray);

                                var ulTypeDataVal = BitConverter.ToUInt64(processArray, 0);

                                Console.WriteLine($"{wdbVars.Fields[f]}(uint64): {ulTypeDataVal}");
                                SharedMethods.WriteToSheet(mainSheet, cellX, cellY, ulTypeDataVal, 5, false);
                                cellY++;

                                strTypeListIndex++;
                                currentRecordDataIndex += 8;
                                break;
                            }

                            var uintDataVal = SharedMethods.DeriveUIntFromSectionData(currentRecordData, currentRecordDataIndex, true);

                            Console.WriteLine($"{wdbVars.Fields[f]}(uint32): {uintDataVal}");
                            SharedMethods.WriteToSheet(mainSheet, cellX, cellY, uintDataVal, 3, false);
                            cellY++;

                            strTypeListIndex++;
                            currentRecordDataIndex += 4;
                            break;
                    }
                }

                Console.WriteLine("");

                strTypeListIndex = 0;
                currentRecordDataIndex = 0;
                sectionPos += 32;
                cellX++;
            }
        }


        private static void ParseRecordsWithoutFields(WDBVariablesXIII wdbVars, IXLWorksheet mainSheet, int cellX, int cellY, BinaryReader br)
        {
            var bitpackedFieldCounter = 0;
            var floatFieldCounter = 0;
            var stringFieldCounter = 0;
            var uintFieldCounter = 0;

            foreach (var strtypelistValue in wdbVars.StrtypelistValues)
            {
                switch (strtypelistValue)
                {
                    case 0:
                        SharedMethods.WriteToSheet(mainSheet, cellX, cellY, $"bitpacked-field_{bitpackedFieldCounter}", 2, true);
                        bitpackedFieldCounter++;
                        break;
                    case 1:
                        SharedMethods.WriteToSheet(mainSheet, cellX, cellY, $"float-field_{floatFieldCounter}", 2, true);
                        floatFieldCounter++;
                        break;
                    case 2:
                        SharedMethods.WriteToSheet(mainSheet, cellX, cellY, $"!!string-field_{stringFieldCounter}", 2, true);
                        stringFieldCounter++;
                        break;
                    case 3:
                        SharedMethods.WriteToSheet(mainSheet, cellX, cellY, $"uint-field_{uintFieldCounter}", 2, true);
                        uintFieldCounter++;
                        break;
                }

                cellY++;
            }


            // Process each record's data
            var sectionPos = br.BaseStream.Position;
            string currentRecordName;
            byte[] currentRecordData;
            int currentRecordDataIndex;
            string bitpackedData;
            float floatValue;
            uint stringValueOffset;
            string stringValue;
            uint uintValue;

            cellX = 2;

            for (int r = 0; r < wdbVars.RecordCount; r++)
            {
                _ = br.BaseStream.Position = sectionPos;
                currentRecordName = br.ReadBytesString(16, false);

                cellY = 1;
                Console.WriteLine($"Record: {currentRecordName}");
                SharedMethods.WriteToSheet(mainSheet, cellX, cellY, currentRecordName, 2, false);
                cellY++;

                currentRecordData = SharedMethods.SaveSectionData(br, false);
                currentRecordDataIndex = 0;

                for (int f = 0; f < wdbVars.FieldCount; f++)
                {
                    switch (wdbVars.StrtypelistValues[f])
                    {
                        case 0:
                            uintValue = SharedMethods.DeriveUIntFromSectionData(currentRecordData, currentRecordDataIndex, true);
                            bitpackedData = uintValue.ToString("X").PadLeft(8, '0');
                            SharedMethods.WriteToSheet(mainSheet, cellX, cellY, "0x" + bitpackedData, 2, false);
                            break;

                        case 1:
                            floatValue = SharedMethods.DeriveFloatFromSectionData(currentRecordData, currentRecordDataIndex, true);
                            SharedMethods.WriteToSheet(mainSheet, cellX, cellY, floatValue, 1, false);
                            break;

                        case 2:
                            stringValueOffset = SharedMethods.DeriveUIntFromSectionData(currentRecordData, currentRecordDataIndex, true); ;
                            stringValue = SharedMethods.DeriveStringFromArray(wdbVars.StringsData, (int)stringValueOffset);

                            if (stringValue == "")
                            {
                                stringValue = "{null}";
                            }
                            SharedMethods.WriteToSheet(mainSheet, cellX, cellY, stringValue, 2, false);
                            break;

                        case 3:
                            uintValue = SharedMethods.DeriveUIntFromSectionData(currentRecordData, currentRecordDataIndex, true);
                            SharedMethods.WriteToSheet(mainSheet, cellX, cellY, uintValue, 3, false);
                            break;
                    }

                    currentRecordDataIndex += 4;
                    cellY++;
                }

                sectionPos += 32;
                cellX++;
            }
        }
    }
}