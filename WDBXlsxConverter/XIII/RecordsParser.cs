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
                WDBMethods.WriteToSheet(mainSheet, cellX, cellY, "Record", 2, true);
                cellY++;

                foreach (var fieldCell in wdbVars.Fields)
                {
                    WDBMethods.WriteToSheet(mainSheet, cellX, cellY, fieldCell, 2, true);
                    cellY++;
                }

                cellX = 2;

                ParseRecordsWithFields(wdbReader, wdbVars, mainSheet, cellX);
            }
            else
            {
                mainSheet = workbook.Worksheets.Add(wdbVars.WDBName);
                WDBMethods.WriteToSheet(mainSheet, cellX, cellY, "Record", 2, true);
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
                WDBMethods.WriteToSheet(mainSheet, cellX, cellY, currentRecordName, 2, false);
                cellY++;

                currentRecordData = WDBMethods.SaveSectionData(wdbReader, false);

                for (int f = 0; f < wdbVars.FieldCount; f++)
                {
                    switch (wdbVars.StrtypelistValues[strTypeListIndex])
                    {
                        case 0:
                            var binaryData = BitOperationHelpers.UIntToBinary(WDBMethods.DeriveUIntFromSectionData(currentRecordData, currentRecordDataIndex, true));
                            var binaryDataIndex = binaryData.Length;
                            var fieldBitsToProcess = 32;

                            int iTypedataVal;
                            uint uTypeDataVal;
                            //float fTypeDataVal;
                            string fTypeBinary;

                            while (fieldBitsToProcess != 0 && f < wdbVars.FieldCount)
                            {
                                var fieldType = wdbVars.Fields[f].Substring(0, 1);
                                var fieldNum = WDBMethods.DeriveFieldNumber(wdbVars.Fields[f]);

                                switch (fieldType)
                                {
                                    // sint
                                    case "i":
                                        if (fieldNum == 0)
                                        {
                                            iTypedataVal = BitOperationHelpers.BinaryToInt(binaryData, binaryDataIndex - 32, 32);
                                            fieldBitsToProcess = 0;

                                            Console.WriteLine($"{wdbVars.Fields[f]}: {iTypedataVal}");
                                            WDBMethods.WriteToSheet(mainSheet, cellX, cellY, iTypedataVal, 4, false);
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
                                            WDBMethods.WriteToSheet(mainSheet, cellX, cellY, iTypedataVal, 4, false);
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
                                            WDBMethods.WriteToSheet(mainSheet, cellX, cellY, uTypeDataVal, 3, false);
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
                                            WDBMethods.WriteToSheet(mainSheet, cellX, cellY, uTypeDataVal, 3, false);
                                            cellY++;

                                            if (fieldBitsToProcess != 0)
                                            {
                                                f++;
                                            }
                                        }
                                        break;

                                    // float (dump as binary) 
                                    case "f":
                                        if (fieldNum == 0)
                                        {
                                            fTypeBinary = binaryData.Substring(binaryDataIndex - 32, 32);
                                            fieldBitsToProcess = 0;

                                            Console.WriteLine($"{wdbVars.Fields[f]}: {fTypeBinary}");
                                            WDBMethods.WriteToSheet(mainSheet, cellX, cellY, fTypeBinary, 2, false);
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

                                            fTypeBinary = binaryData.Substring(binaryDataIndex, fieldNum);
                                            fieldBitsToProcess -= fieldNum;

                                            Console.WriteLine($"{wdbVars.Fields[f]}: {fTypeBinary}");
                                            WDBMethods.WriteToSheet(mainSheet, cellX, cellY, fTypeBinary, 2, false);
                                            cellY++;

                                            if (fieldBitsToProcess != 0)
                                            {
                                                f++;
                                            }
                                        }
                                        break;

                                        // // float 
                                        //case "f":
                                        //    if (fieldNum == 0)
                                        //    {
                                        //        fTypeDataVal = BitOperationHelpers.BinaryToFloat(binaryData, binaryDataIndex - 32, 32);
                                        //        fieldBitsToProcess = 0;

                                        //        Console.WriteLine($"{wdbVars.Fields[f]}: {fTypeDataVal}");
                                        //        WDBMethods.WriteToWorkSheet(mainSheet, recordCellX, recordCellY, fTypeDataVal.ToString(), false);
                                        //        recordCellY++;

                                        //        break;
                                        //    }
                                        //    if (fieldNum > fieldBitsToProcess)
                                        //    {
                                        //        f--;
                                        //        fieldBitsToProcess = 0;
                                        //        continue;
                                        //    }
                                        //    else
                                        //    {
                                        //        binaryDataIndex -= fieldNum;

                                        //        fTypeDataVal = BitOperationHelpers.BinaryToFloat(binaryData, binaryDataIndex, fieldNum);
                                        //        fieldBitsToProcess -= fieldNum;

                                        //        Console.WriteLine($"{wdbVars.Fields[f]}: {fTypeDataVal}");
                                        //        WDBMethods.WriteToWorkSheet(mainSheet, recordCellX, recordCellY, fTypeDataVal.ToString(), false);
                                        //        recordCellY++;

                                        //        if (fieldBitsToProcess != 0)
                                        //        {
                                        //            f++;
                                        //        }
                                        //    }
                                        //    break;
                                }
                            }

                            strTypeListIndex++;
                            currentRecordDataIndex += 4;
                            break;

                        // float value
                        case 1:
                            var floatDataVal = WDBMethods.DeriveFloatFromSectionData(currentRecordData, currentRecordDataIndex, true);

                            Console.WriteLine($"{wdbVars.Fields[f]}: {floatDataVal}");
                            WDBMethods.WriteToSheet(mainSheet, cellX, cellY, floatDataVal, 1, false);
                            cellY++;

                            strTypeListIndex++;
                            currentRecordDataIndex += 4;
                            break;

                        // !!string section offset
                        case 2:
                            var stringDataOffset = WDBMethods.DeriveUIntFromSectionData(currentRecordData, currentRecordDataIndex, true);
                            var derivedString = WDBMethods.DeriveStringFromArray(wdbVars.StringsData, (int)stringDataOffset);

                            if (derivedString == "")
                            {
                                derivedString = "{null}";
                            }

                            Console.WriteLine($"{wdbVars.Fields[f]}: {derivedString}");
                            WDBMethods.WriteToSheet(mainSheet, cellX, cellY, derivedString, 2, false);
                            cellY++;

                            strTypeListIndex++;
                            currentRecordDataIndex += 4;
                            break;

                        // uint
                        case 3:
                            if (wdbVars.Fields[f].StartsWith("u64"))
                            {
                                var processArray = new byte[8];
                                Array.ConstrainedCopy(currentRecordData, currentRecordDataIndex, processArray, 0, 8);
                                Array.Reverse(processArray);

                                var ulTypeDataVal = BitConverter.ToUInt64(processArray, 0);

                                Console.WriteLine($"{wdbVars.Fields[f]}(uint64): {ulTypeDataVal}");
                                WDBMethods.WriteToSheet(mainSheet, cellX, cellY, ulTypeDataVal, 5, false);
                                cellY++;

                                strTypeListIndex++;
                                currentRecordDataIndex += 8;
                                break;
                            }

                            var uintDataVal = WDBMethods.DeriveUIntFromSectionData(currentRecordData, currentRecordDataIndex, true);

                            Console.WriteLine($"{wdbVars.Fields[f]}(uint32): {uintDataVal}");
                            WDBMethods.WriteToSheet(mainSheet, cellX, cellY, uintDataVal, 3, false);
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
            foreach (var strtypelistValue in wdbVars.StrtypelistValues)
            {
                switch (strtypelistValue)
                {
                    case 0:
                        WDBMethods.WriteToSheet(mainSheet, cellX, cellY, "BitPacked", 2, true);
                        break;
                    case 1:
                        WDBMethods.WriteToSheet(mainSheet, cellX, cellY, "Float", 2, true);
                        break;
                    case 2:
                        WDBMethods.WriteToSheet(mainSheet, cellX, cellY, "!!string", 2, true);
                        break;
                    case 3:
                        WDBMethods.WriteToSheet(mainSheet, cellX, cellY, "Unsigned Integer", 2, true);
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
                WDBMethods.WriteToSheet(mainSheet, cellX, cellY, currentRecordName, 2, false);
                cellY++;

                currentRecordData = WDBMethods.SaveSectionData(br, false);
                currentRecordDataIndex = 0;

                for (int f = 0; f < wdbVars.FieldCount; f++)
                {
                    switch (wdbVars.StrtypelistValues[f])
                    {
                        case 0:
                            uintValue = WDBMethods.DeriveUIntFromSectionData(currentRecordData, currentRecordDataIndex, true);
                            bitpackedData = uintValue.ToString("X").PadLeft(8, '0');
                            WDBMethods.WriteToSheet(mainSheet, cellX, cellY, "0x" + bitpackedData, 2, false);
                            break;

                        case 1:
                            floatValue = WDBMethods.DeriveFloatFromSectionData(currentRecordData, currentRecordDataIndex, true);
                            WDBMethods.WriteToSheet(mainSheet, cellX, cellY, floatValue, 1, false);
                            break;

                        case 2:
                            stringValueOffset = WDBMethods.DeriveUIntFromSectionData(currentRecordData, currentRecordDataIndex, true); ;
                            stringValue = WDBMethods.DeriveStringFromArray(wdbVars.StringsData, (int)stringValueOffset);

                            if (stringValue == "")
                            {
                                stringValue = "{null}";
                            }
                            WDBMethods.WriteToSheet(mainSheet, cellX, cellY, stringValue, 2, false);
                            break;

                        case 3:
                            uintValue = WDBMethods.DeriveUIntFromSectionData(currentRecordData, currentRecordDataIndex, true);
                            WDBMethods.WriteToSheet(mainSheet, cellX, cellY, uintValue, 3, false);
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