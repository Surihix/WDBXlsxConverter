using ClosedXML.Excel;
using System.IO;
using System.Text;

namespace WDBXlsxConverter.XIII
{
    internal class SectionsParser
    {
        public static void MainSections(BinaryReader wdbReader, WDBVariablesXIII wdbVars)
        {
            // Parse main sections
            long currentSectionNamePos = 16;
            string sectioNameRead;

            while (true)
            {
                wdbReader.BaseStream.Position = currentSectionNamePos;
                sectioNameRead = wdbReader.ReadBytesString(16, false);

                // Break the loop if its
                // not a valid "!" section
                if (!sectioNameRead.StartsWith("!!"))
                {
                    _ = wdbReader.BaseStream.Position = currentSectionNamePos;
                    break;
                }

                // !!sheetname check
                if (sectioNameRead == "!!sheetname")
                {
                    WDBMethods.ErrorExit("Specified WDB file is from XIII-2 or LR. set the gamecode to -ff132 to convert this file.");
                }

                // !!strArray check
                if (sectioNameRead == "!!strArray")
                {
                    WDBMethods.ErrorExit("Specified WDB file is from XIII-2 or LR. set the gamecode to -ff132 to convert this file.");
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
                    wdbVars.StrtypelistData = WDBMethods.SaveSectionData(wdbReader, false);

                    if (wdbVars.StrtypelistData.Length != 0)
                    {
                        wdbVars.StrtypelistValues = WDBMethods.GetSectionDataValues(wdbVars.StrtypelistData);
                        wdbVars.FieldCount = (uint)wdbVars.StrtypelistValues.Count;
                    }

                    wdbVars.RecordCount--;
                }

                // !!typelist
                if (sectioNameRead == wdbVars.TypelistSectionName)
                {
                    wdbVars.TypelistData = WDBMethods.SaveSectionData(wdbReader, false);

                    if (wdbVars.TypelistData.Length != 0)
                    {
                        wdbVars.TypelistValues = WDBMethods.GetSectionDataValues(wdbVars.TypelistData);
                    }

                    wdbVars.RecordCount--;
                }

                // !!version
                if (sectioNameRead == wdbVars.VersionSectionName)
                {
                    wdbVars.VersionData = WDBMethods.SaveSectionData(wdbReader, false);
                    wdbVars.RecordCount--;
                }

                currentSectionNamePos += 32;
            }

            // Check if the !!strtypelist
            // is parsed 
            if (wdbVars.StrtypelistData.Length == 0)
            {
                WDBMethods.ErrorExit("!!strtypelist section was not present in the file.");
            }
        }


        public static void MainSectionsToWB(WDBVariablesXIII wdbVars, XLWorkbook workbook)
        {
            int cellX = 1;  // vertical
            int cellY = 1;  // horizontal

            if (WDBDictsXIII.RecordIDs.ContainsKey(wdbVars.WDBName) && !wdbVars.IgnoreKnown)
            {
                wdbVars.IsKnown = true;
                wdbVars.SheetName = WDBDictsXIII.RecordIDs[wdbVars.WDBName];
                wdbVars.FieldCount = (uint)WDBDictsXIII.FieldNames[wdbVars.SheetName].Count;

                wdbVars.Fields = new string[wdbVars.FieldCount];

                // Write all of the field names 
                // if the file is fully known
                for (int sf = 0; sf < wdbVars.FieldCount; sf++)
                {
                    var derivedString = WDBDictsXIII.FieldNames[wdbVars.SheetName][sf];
                    wdbVars.Fields[sf] = derivedString;
                }
            }


            // Write all the stringData in the
            // worksheet
            if (wdbVars.HasStringSection)
            {
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

                    WDBMethods.WriteToSheet(stringSheet, cellX, cellY, stringStartPos.ToString(), 3, false);
                    WDBMethods.WriteToSheet(stringSheet, cellX, cellY + 1, derivedString, 2, false);
                    WDBMethods.WriteToSheet(stringSheet, cellX, cellY + 2, derivedStringLength.ToString(), 3, false);

                    stringStartPos += derivedStringLength;

                    cellX++;
                }

                stringSheet.Rows().AdjustToContents();
                stringSheet.Columns().AdjustToContents();
            }


            // Parse and write strtypelistData
            // section in the worksheet
            //
            // Depending on whether the file is
            // fully understood,
            // either write only the values or
            // the values with the field names
            cellX = 1;
            var strtypelistSheet = workbook.Worksheets.Add(wdbVars.StrtypelistSectionName);

            if (wdbVars.IsKnown)
            {
                WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY, "Field Name", 2, true);
                WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, "Strtypelist Value", 2, true);
                
                cellX++;

                var strTypeListIndex = 0;

                for (int f = 0; f < wdbVars.FieldCount; f++)
                {
                    var currentField = wdbVars.Fields[f];
                    var strtypelistValue = wdbVars.StrtypelistValues[strTypeListIndex];

                    switch (strtypelistValue)
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
                                            WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistValue, 4, false);
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
                                            WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistValue, 4, false);
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
                                            WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistValue, 4, false);
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
                                            WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistValue, 4, false);
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
                                            WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistValue, 4, false);
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
                                            WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistValue, 4, false);
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

                            strTypeListIndex++;
                            break;

                        // float value
                        case 1:
                            WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY, $"{currentField}", 2, false);
                            WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistValue, 4, false);
                            cellX++;

                            strTypeListIndex++;
                            break;

                        // !!string section offset
                        case 2:
                            WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY, $"{currentField}", 2, false);
                            WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistValue, 4, false);
                            cellX++;

                            strTypeListIndex++;
                            break;

                        // uint
                        case 3:
                            WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY, $"{currentField}", 2, false);
                            WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY + 1, strtypelistValue, 4, false);
                            cellX++;

                            strTypeListIndex++;
                            break;
                    }
                }
            }
            else
            {
                WDBMethods.WriteToSheet(strtypelistSheet, cellX, cellY, "Strtypelist Value", 2, true);
                cellX++;

                WDBMethods.WriteValuesListToSheet(wdbVars.StrtypelistValues, strtypelistSheet, cellX, cellY);
            }

            strtypelistSheet.Rows().AdjustToContents();
            strtypelistSheet.Columns().AdjustToContents();


            // Parse and write typelistData
            // section in the worksheet
            cellX = 1;
            var typelistSheet = workbook.Worksheets.Add(wdbVars.TypelistSectionName);
            WDBMethods.WriteToSheet(typelistSheet, cellX, cellY, "Typelist Value", 2, true);

            cellX++;

            WDBMethods.WriteValuesListToSheet(wdbVars.TypelistValues, typelistSheet, cellX, cellY);

            typelistSheet.Rows().AdjustToContents();
            typelistSheet.Columns().AdjustToContents();
        }
    }
}