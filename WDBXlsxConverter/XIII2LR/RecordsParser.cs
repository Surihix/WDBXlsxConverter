using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace WDBXlsxConverter.XIII2LR
{
    internal class RecordsParser
    {
        public static void ProcessRecords(BinaryReader wdbReader, WDBVariablesXIII2LR wdbVars, XLWorkbook workbook)
        {
            // Process each record's data
            var sectionPos = wdbReader.BaseStream.Position;
            string currentRecordName;
            byte[] currentRecordData;
            var strtypelistIndex = 0;
            var currentRecordDataIndex = 0;

            var cellX = 1;  // vertical
            var cellY = 1;  // horizontal


            // Write all the fields in the 
            // worksheet
            var mainSheet = workbook.Worksheets.Add(wdbVars.SheetName);
            SharedMethods.WriteToSheet(mainSheet, cellX, cellY, "Record", 2, true);
            cellY++;

            foreach (var fieldCell in wdbVars.Fields)
            {
                SharedMethods.WriteToSheet(mainSheet, cellX, cellY, fieldCell, 2, true);
                cellY++;
            }

            cellX = 2;

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
                    switch (wdbVars.StrtypelistValues[strtypelistIndex])
                    {
                        case 0:
                            var binaryData = BitOperationHelpers.UIntToBinary(SharedMethods.DeriveUIntFromSectionData(currentRecordData, currentRecordDataIndex, true));
                            var binaryDataIndex = binaryData.Length;
                            var fieldBitsToProcess = 32;

                            int iTypedataVal;
                            uint uTypeDataVal;
                            int fTypeDataVal;

                            uint strArrayTypeDataVal;
                            string strArrayTypeDictKey;
                            List<string> strArrayTypeDictList;

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

                                    // (s#) strArray item index
                                    case "s":
                                        if (fieldNum > fieldBitsToProcess)
                                        {
                                            f--;
                                            fieldBitsToProcess = 0;
                                            continue;
                                        }
                                        else
                                        {
                                            binaryDataIndex -= fieldNum;

                                            strArrayTypeDataVal = BitOperationHelpers.BinaryToUInt(binaryData, binaryDataIndex, fieldNum);
                                            fieldBitsToProcess -= fieldNum;

                                            strArrayTypeDictKey = wdbVars.Fields[f];
                                            strArrayTypeDictList = wdbVars.StrArrayDict[strArrayTypeDictKey];

                                            Console.WriteLine($"{strArrayTypeDictKey}: {strArrayTypeDictList[(int)strArrayTypeDataVal]}");
                                            SharedMethods.WriteToSheet(mainSheet, cellX, cellY, strArrayTypeDictList[(int)strArrayTypeDataVal], 2, false);
                                            cellY++;

                                            if (fieldBitsToProcess != 0)
                                            {
                                                f++;
                                            }
                                        }
                                        break;
                                }
                            }
                            
                            strtypelistIndex++;
                            currentRecordDataIndex += 4;
                            break;

                        // float value
                        case 1:
                            var floatDataVal = SharedMethods.DeriveFloatFromSectionData(currentRecordData, currentRecordDataIndex, true);

                            Console.WriteLine($"{wdbVars.Fields[f]}: {floatDataVal}");
                            SharedMethods.WriteToSheet(mainSheet, cellX, cellY, floatDataVal, 1, false);
                            cellY++;

                            strtypelistIndex++;
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

                            strtypelistIndex++;
                            currentRecordDataIndex += 4;
                            break;

                        // uint
                        case 3:
                            var uintDataVal = SharedMethods.DeriveUIntFromSectionData(currentRecordData, currentRecordDataIndex, true);

                            Console.WriteLine($"{wdbVars.Fields[f]}(uint32): {uintDataVal}");
                            SharedMethods.WriteToSheet(mainSheet, cellX, cellY, uintDataVal, 3, false);
                            cellY++;

                            strtypelistIndex++;
                            currentRecordDataIndex += 4;
                            break;
                    }
                }

                Console.WriteLine("");

                strtypelistIndex = 0;
                currentRecordDataIndex = 0;
                sectionPos += 32;
                cellX++;
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
    }
}