using ClosedXML.Excel;
using System;
using System.IO;
using System.Text;

namespace WDBXlsxConverter.XIII2LR
{
    internal class SectionsParser
    {
        public static void MainSections(BinaryReader wdbReader, WDBVariablesXIII2LR wdbVars)
        {
            // Parse main sections
            long currentSectionNamePos = 16;
            string sectioNameRead;

            wdbVars.StrtypelistData = new byte[] { };
            wdbVars.StructItemData = new byte[] { };
            wdbVars.FieldCount = 0;


            while (true)
            {
                wdbReader.BaseStream.Position = currentSectionNamePos;
                sectioNameRead = wdbReader.ReadBytesString(16, false);

                // Break the loop if its
                // not a valid "!" section
                if (!sectioNameRead.StartsWith("!"))
                {
                    if (wdbVars.SheetName == "")
                    {
                        wdbVars.SheetName = wdbVars.SheetNameSectionName;

                        Console.WriteLine("");
                        Console.WriteLine("");
                        Console.WriteLine($"{wdbVars.SheetNameSectionName}: {wdbVars.SheetName}");
                        Console.WriteLine("");
                        Console.WriteLine("");
                    }

                    _ = wdbReader.BaseStream.Position = currentSectionNamePos;
                    break;
                }

                // !!sheetname
                if (sectioNameRead == wdbVars.SheetNameSectionName)
                {
                    _ = wdbReader.BaseStream.Position = wdbReader.ReadBytesUInt32(true);
                    wdbVars.SheetName = wdbReader.ReadStringTillNull();
                    wdbVars.RecordCount--;
                }

                // !!strArray
                if (sectioNameRead == wdbVars.StrArraySectionName)
                {
                    wdbVars.HasStrArraySection = true;

                    _ = wdbReader.BaseStream.Position = currentSectionNamePos;
                    StrArrayParser.SubSections(wdbReader, wdbVars);
                }

                // !!string
                if (sectioNameRead == wdbVars.StringSectionName)
                {
                    wdbVars.HasStringSection = true;

                    wdbVars.StringsData = WDBMethods.SaveSectionData(wdbReader, false);
                    wdbVars.RecordCount--;
                }

                // !!strtypelist
                if (sectioNameRead == wdbVars.StrtypelistSectionName)
                {
                    wdbVars.parseStrtypelistAsV1 = true;
                    wdbVars.StrtypelistData = WDBMethods.SaveSectionData(wdbReader, false);
                    wdbVars.RecordCount--;
                }

                // !!strtypelistb
                if (sectioNameRead == wdbVars.StrtypelistbSectionName)
                {
                    wdbVars.parseStrtypelistAsV1 = false;
                    wdbVars.StrtypelistData = WDBMethods.SaveSectionData(wdbReader, false);
                    wdbVars.RecordCount--;
                }

                // !!typelist
                if (sectioNameRead == wdbVars.TypelistSectionName)
                {
                    wdbVars.hasTypelistSection = true;
                    wdbVars.TypelistData = WDBMethods.SaveSectionData(wdbReader, false);
                    wdbVars.RecordCount--;
                }

                // !!version
                if (sectioNameRead == wdbVars.VersionSectionName)
                {
                    wdbVars.VersionData = WDBMethods.SaveSectionData(wdbReader, false);
                    wdbVars.RecordCount--;
                }

                // !structitem
                if (sectioNameRead == wdbVars.StructItemSectionName)
                {
                    wdbVars.StructItemData = WDBMethods.SaveSectionData(wdbReader, false);
                    wdbVars.RecordCount--;
                }

                // !structitemnum
                if (sectioNameRead == wdbVars.StructItemNumSectionName)
                {
                    wdbVars.FieldCount = BitConverter.ToUInt32(WDBMethods.SaveSectionData(wdbReader, true), 0);
                    wdbVars.RecordCount--;
                }

                currentSectionNamePos += 32;
            }


            // Check if the important 
            // sections are all parsed
            var imptSectionsParsed = wdbVars.StrtypelistData.Length != 0 && wdbVars.StructItemData.Length != 0 && wdbVars.FieldCount != 0;

            if (!imptSectionsParsed)
            {
                WDBMethods.ErrorExit("Necessary sections were unable to be processed correctly.");
            }

            if (wdbVars.SheetName == "" || wdbVars.SheetName == null)
            {
                wdbVars.SheetName = wdbVars.WDBName;
            }

            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine($"{wdbVars.SheetNameSectionName}: {wdbVars.SheetName}");
            Console.WriteLine("");
            Console.WriteLine("");


            // Process !structitem data
            wdbVars.Fields = new string[wdbVars.FieldCount];
            var stringStartPos = 0;

            for (int sf = 0; sf < wdbVars.FieldCount; sf++)
            {
                var derivedString = WDBMethods.DeriveStringFromArray(wdbVars.StructItemData, stringStartPos);

                if (derivedString == "")
                {
                    wdbVars.Fields[sf] = "{null}";
                }
                else
                {
                    wdbVars.Fields[sf] = derivedString;
                }

                stringStartPos += Encoding.UTF8.GetByteCount(derivedString) + 1;
            }


            // Process strArray sections
            // data
            if (wdbVars.HasStrArraySection)
            {
                Console.WriteLine($"Organizing {wdbVars.StrArraySectionName} data....");

                StrArrayParser.ArrangeArrayData(wdbVars);

                Console.WriteLine("");
                Console.WriteLine("");
            }
        }


        public static void MainSectionsToWB(WDBVariablesXIII2LR wdbVars, XLWorkbook workbook)
        {
            var cellX = 1;  // vertical
            var cellY = 1;  // horizontal

            // Write all the strArrayData in the
            // worksheet
            if (wdbVars.HasStrArraySection)
            {
                var strArraySheet = workbook.Worksheets.Add(wdbVars.StrArraySectionName);
                var strArrayNoIterator = 0;

                foreach (var k in wdbVars.StrArrayDict)
                {
                    WDBMethods.WriteToSheet(strArraySheet, cellX, cellY, $"(Array {strArrayNoIterator}) Index", 2, true);
                    WDBMethods.WriteToSheet(strArraySheet, cellX, cellY + 1, $"{k.Key}", 2, true);

                    var prevCellY = cellY;

                    cellY += 3;
                    cellX++;

                    var indexIterator = 0;

                    foreach (var l in wdbVars.StrArrayDict[k.Key])
                    {
                        WDBMethods.WriteToSheet(strArraySheet, cellX, prevCellY, indexIterator, 4, false);
                        indexIterator++;

                        WDBMethods.WriteToSheet(strArraySheet, cellX, prevCellY + 1, l, 2, false);
                        cellX++;
                    }

                    cellX = 1;
                    strArrayNoIterator++;
                }

                strArraySheet.Rows().AdjustToContents();
                strArraySheet.Columns().AdjustToContents();
            }


            // Write all the stringData in the
            // worksheet
            if (wdbVars.HasStringSection)
            {
                cellX = 1;  // vertical
                cellY = 1;  // horizontal

                var stringSheet = workbook.Worksheets.Add(wdbVars.StringSectionName);
                WDBMethods.WriteToSheet(stringSheet, cellX, cellY, "Position", 2, true);
                WDBMethods.WriteToSheet(stringSheet, cellX, cellY + 1, "String", 2, true);
                WDBMethods.WriteToSheet(stringSheet, cellX, cellY + 2, "Length (with null byte)", 2, true);

                cellX++;

                var stringStartPos = 0;
                while (true)
                {
                    if (stringStartPos >= wdbVars.StringsData.Length)
                    {
                        break;
                    }

                    var derivedString = WDBMethods.DeriveStringFromArray(wdbVars.StringsData, stringStartPos);
                    var derivedStringLength = Encoding.UTF8.GetByteCount(derivedString) + 1;

                    if (derivedString == "")
                    {
                        derivedString = "{null}";
                    }

                    WDBMethods.WriteToSheet(stringSheet, cellX, cellY, stringStartPos, 4, false);
                    WDBMethods.WriteToSheet(stringSheet, cellX, cellY + 1, derivedString, 2, false);
                    WDBMethods.WriteToSheet(stringSheet, cellX, cellY + 2, derivedStringLength, 4, false);

                    stringStartPos += derivedStringLength;

                    cellX++;
                }

                stringSheet.Rows().AdjustToContents();
                stringSheet.Columns().AdjustToContents();
            }


            // Parse and write the strtypelistData
            // section in the worksheet
            cellX = 1;
            cellY = 1;
            var strtypelistSheetName = wdbVars.parseStrtypelistAsV1 ? wdbVars.StrtypelistSectionName : wdbVars.StrtypelistbSectionName;
            var strtypelistSheet = workbook.Worksheets.Add(strtypelistSheetName);
            var strtypelistValueFieldName = wdbVars.parseStrtypelistAsV1 ? "Strtypelist Value" : "Strtypelistb Value";

            WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY, "Field Name", 2, true);
            WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistValueFieldName, 2, true);

            cellX++;

            var strtypelistbIndex = 0;
            var currentStrtypelistData = new byte[4];
            var strtypelistIndexAdjust = wdbVars.parseStrtypelistAsV1 ? 4 : 1;
            int strtypelistbValue;

            for (int f = 0; f < wdbVars.FieldCount; f++)
            {
                var currentField = wdbVars.Fields[f];

                if (wdbVars.parseStrtypelistAsV1)
                {
                    Array.ConstrainedCopy(wdbVars.StrtypelistData, strtypelistbIndex, currentStrtypelistData, 0, 4);
                    Array.Reverse(currentStrtypelistData);
                    strtypelistbValue = (int)BitConverter.ToUInt32(currentStrtypelistData, 0);
                }
                else
                {
                    strtypelistbValue = wdbVars.StrtypelistData[strtypelistbIndex];
                }

                wdbVars.StrtypelistValues.Add(strtypelistbValue);

                switch (strtypelistbValue)
                {
                    case 0:
                        var fieldBitsToProcess = 32;

                        while (fieldBitsToProcess != 0 && f < wdbVars.FieldCount)
                        {
                            currentField = wdbVars.Fields[f];
                            var fieldType = currentField.Substring(0, 1);
                            var fieldNum = WDBMethods.DeriveFieldNumber(currentField);

                            switch (fieldType)
                            {
                                // sint
                                case "i":
                                    if (fieldNum == 0)
                                    {
                                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY, $"{currentField}", 2, false);
                                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistbValue, 4, false);
                                        cellX++;

                                        fieldBitsToProcess = 0;
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
                                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY, $"{currentField}", 2, false);
                                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistbValue, 4, false);
                                        cellX++;

                                        fieldBitsToProcess -= fieldNum;

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
                                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY, $"{currentField}", 2, false);
                                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistbValue, 4, false);
                                        cellX++;

                                        fieldBitsToProcess = 0;
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
                                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY, $"{currentField}", 2, false);
                                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistbValue, 4, false);
                                        cellX++;

                                        fieldBitsToProcess -= fieldNum;

                                        if (fieldBitsToProcess != 0)
                                        {
                                            f++;
                                        }
                                    }
                                    break;

                                // float
                                case "f":
                                    if (fieldNum == 0)
                                    {
                                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY, $"{currentField}", 2, false);
                                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistbValue, 4, false);
                                        cellX++;

                                        fieldBitsToProcess = 0;
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
                                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY, $"{currentField}", 2, false);
                                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistbValue, 4, false);
                                        cellX++;

                                        fieldBitsToProcess -= fieldNum;

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
                                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY, $"{currentField}", 2, false);
                                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistbValue, 4, false);
                                        cellX++;

                                        fieldBitsToProcess -= fieldNum;

                                        if (fieldBitsToProcess != 0)
                                        {
                                            f++;
                                        }
                                    }
                                    break;
                            }
                        }

                        strtypelistbIndex += strtypelistIndexAdjust;
                        break;

                    // float value
                    case 1:
                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY, $"{currentField}", 2, false);
                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistbValue, 4, false);
                        cellX++;

                        strtypelistbIndex += strtypelistIndexAdjust;
                        break;

                    // !!string section offset
                    case 2:
                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY, $"{currentField}", 2, false);
                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistbValue, 4, false);
                        cellX++;

                        strtypelistbIndex += strtypelistIndexAdjust;
                        break;

                    // uint
                    case 3:
                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY, $"{currentField}", 2, false);
                        WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistbValue, 4, false);
                        cellX++;

                        strtypelistbIndex += strtypelistIndexAdjust;
                        break;
                }
            }

            strtypelistSheet.Rows().AdjustToContents();
            strtypelistSheet.Columns().AdjustToContents();


            // Write all the typelist data in the
            // worksheet
            if (wdbVars.hasTypelistSection)
            {
                cellX = 1;
                cellY = 1;

                var typelistSheet = workbook.Worksheets.Add(wdbVars.TypelistSectionName);
                WDBMethods.WriteToSheet(typelistSheet, cellX, cellY, "Typelist Value", 2, true);

                cellX++;

                var sectionIndex = 0;
                for (int i = 0; i < wdbVars.TypelistData.Length / 4; i++)
                {
                    WDBMethods.WriteToSheet(typelistSheet, cellX, cellY, WDBMethods.DeriveUIntFromSectionData(wdbVars.TypelistData, sectionIndex, true), 4, false);
                    cellX++;

                    sectionIndex += 4;
                }

                typelistSheet.Rows().AdjustToContents();
                typelistSheet.Columns().AdjustToContents();
            }
        }
    }
}