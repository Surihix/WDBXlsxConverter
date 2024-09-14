using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace WDBXlsxConverter.XIII2LR
{
    internal class RecordsParser
    {
        public static void ProcessRecords(BinaryReader br, WDBVariablesXIII2LR wdbVars, XLWorkbook workbook)
        {
            // Process each record's data
            var sectionPos = br.BaseStream.Position;
            string currentRecordName;
            byte[] currentRecordData;
            var strtypelistIndex = 0;
            var currentRecordDataIndex = 0;

            var cellX = 1;  // vertical
            var cellY = 1;  // horizontal


            // Write all the fields in the 
            // worksheet
            var mainSheet = workbook.Worksheets.Add(wdbVars.SheetName);
            WDBMethods.WriteToSheet(mainSheet, cellX, cellY, "Record", 2, true);
            cellY++;

            foreach (var fieldCell in wdbVars.Fields)
            {
                WDBMethods.WriteToSheet(mainSheet, cellX, cellY, fieldCell, 2, true);
                cellY++;
            }

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

                for (int f = 0; f < wdbVars.FieldCount; f++)
                {
                    switch (wdbVars.StrtypelistValues[strtypelistIndex])
                    {
                        case 0:
                            var binaryData = BitOperationHelpers.UIntToBinary(WDBMethods.DeriveUIntFromSectionData(currentRecordData, currentRecordDataIndex, true));
                            var binaryDataIndex = binaryData.Length;
                            var fieldBitsToProcess = 32;

                            int iTypedataVal;
                            uint uTypeDataVal;
                            //float fTypeDataVal;
                            string fTypeBinary;
                            uint strArrayTypeDataVal;
                            string strArrayTypeDictKey;
                            List<string> strArrayTypeDictList;

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
                                            WDBMethods.WriteToSheet(mainSheet, cellX, cellY, strArrayTypeDictList[(int)strArrayTypeDataVal], 2, false);
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
                            var floatDataVal = WDBMethods.DeriveFloatFromSectionData(currentRecordData, currentRecordDataIndex, true);

                            Console.WriteLine($"{wdbVars.Fields[f]}: {floatDataVal}");
                            WDBMethods.WriteToSheet(mainSheet, cellX, cellY, floatDataVal, 1, false);
                            cellY++;

                            strtypelistIndex++;
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

                            strtypelistIndex++;
                            currentRecordDataIndex += 4;
                            break;

                        // uint
                        case 3:
                            var uintDataVal = WDBMethods.DeriveUIntFromSectionData(currentRecordData, currentRecordDataIndex, true);

                            Console.WriteLine($"{wdbVars.Fields[f]}(uint32): {uintDataVal}");
                            WDBMethods.WriteToSheet(mainSheet, cellX, cellY, uintDataVal, 3, false);
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