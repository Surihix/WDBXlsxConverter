using System;
using System.Collections.Generic;
using System.IO;

namespace WDBXlsxConverter.XIII2LR
{
    internal class StrArrayParser
    {
        public static void SubSections(BinaryReader br, WDBVariablesXIII2LR wdbVars)
        {
            var readPos = br.BaseStream.Position;

            // !!strArray
            _ = br.BaseStream.Position = readPos;
            WDBMethods.CheckSectionName(br, wdbVars.StrArraySectionName);
            wdbVars.StrArrayData = WDBMethods.SaveSectionData(br, false);
            wdbVars.RecordCount--;

            // !!strArrayInfo
            _ = br.BaseStream.Position = readPos + 32;
            WDBMethods.CheckSectionName(br, wdbVars.StrArrayInfoSectionName);

            _ = br.BaseStream.Position = br.ReadBytesUInt32(true) + 2;
            wdbVars.OffsetsPerValue = br.ReadByte();
            wdbVars.BitsPerOffset = br.ReadByte();
            wdbVars.RecordCount--;

            Console.WriteLine("[StrArray]");
            Console.WriteLine($"Offsets per value: {wdbVars.OffsetsPerValue}");
            Console.WriteLine($"Bits per offset: {wdbVars.BitsPerOffset}");
            Console.WriteLine("");


            // !!strArrayList
            _ = br.BaseStream.Position = readPos + 64;
            WDBMethods.CheckSectionName(br, wdbVars.StrArrayListSectionName);
            wdbVars.StrArrayListData = WDBMethods.SaveSectionData(br, false);
            wdbVars.RecordCount--;
        }



        public static void ArrangeArrayData(WDBVariablesXIII2LR wdbVars)
        {
            // Collect all !!strArray Offsets
            byte[] tmpReadArray;

            for (int a = 0; a < wdbVars.StrArrayListData.Length; a += 4)
            {
                tmpReadArray = new byte[4];
                Array.ConstrainedCopy(wdbVars.StrArrayListData, a, tmpReadArray, 0, 4);

                Array.Reverse(tmpReadArray);
                wdbVars.StrArrayOffsets.Add(BitConverter.ToUInt32(tmpReadArray, 0));
            }


            // Process numbered s# fields in
            // !structitem data
            int fieldNumber;

            foreach (var fieldItem in wdbVars.Fields)
            {
                if (fieldItem.StartsWith("s"))
                {
                    fieldNumber = WDBMethods.DeriveFieldNumber(fieldItem);

                    if (fieldNumber != 0)
                    {
                        wdbVars.NumStringFields.Add(fieldItem);
                    }
                }
            }

            if (wdbVars.StrArrayOffsets.Count != wdbVars.NumStringFields.Count)
            {
                WDBMethods.ErrorExit("StrArrayOffsets count does not match with the detected amount of s# fields");
            }


            // Copy !!strArray data into a stream and 
            // open a reader for this stream
            using (var strArrayStream = new MemoryStream())
            {
                using (var strArrayReader = new BinaryReader(strArrayStream))
                {
                    var strArrayDataLen = wdbVars.StrArrayData.Length;
                    strArrayStream.Write(wdbVars.StrArrayData, 0, strArrayDataLen);
                    strArrayStream.Seek(0, SeekOrigin.Begin);

                    // Copy !!strings data into a stream and 
                    // open a reader for this stream
                    using (var stringsStream = new MemoryStream())
                    {
                        using (var stringsReader = new BinaryReader(stringsStream))
                        {
                            stringsStream.Write(wdbVars.StringsData, 0, wdbVars.StringsData.Length);
                            stringsStream.Seek(0, SeekOrigin.Begin);


                            uint strArrayBinaryUInt = 0;
                            var strArrayBinary = "";
                            uint stringOffset = 0;
                            var binaryReadPos = 0;
                            var arrayIterator = 0;
                            var currentString = "";
                            uint currentArrayItemIndex = 0;
                            bool buildArray = true;

                            foreach (var listOffset in wdbVars.StrArrayOffsets)
                            {
                                wdbVars.ProcessStringsList = new List<string>();

                                while (buildArray)
                                {
                                    // Get the value with which the string
                                    // offset values will be derived and
                                    // convert the value to binary
                                    //
                                    // 'binaryReadPos' value will be read from
                                    // right to left
                                    strArrayBinaryUInt = strArrayReader.ReadBytesUInt32(true);
                                    strArrayBinary = BitOperationHelpers.UIntToBinary(strArrayBinaryUInt);
                                    binaryReadPos = strArrayBinary.Length - wdbVars.BitsPerOffset;

                                    // Get each offset's values
                                    for (int o = 0; o < wdbVars.OffsetsPerValue; o++)
                                    {
                                        stringOffset = BitOperationHelpers.BinaryToUInt(strArrayBinary, binaryReadPos, wdbVars.BitsPerOffset);

                                        stringsReader.BaseStream.Position = stringOffset;
                                        currentString = stringsReader.ReadStringTillNull();

                                        if (currentString == "")
                                        {
                                            currentString = "{null}";
                                        }

                                        wdbVars.ProcessStringsList.Add(currentString);

                                        currentArrayItemIndex++;
                                        binaryReadPos -= wdbVars.BitsPerOffset;
                                    }

                                    if (strArrayReader.BaseStream.Position == strArrayDataLen)
                                    {
                                        buildArray = false;
                                    }

                                    if (wdbVars.StrArrayOffsets.Contains((uint)strArrayReader.BaseStream.Position))
                                    {
                                        buildArray = false;
                                    }
                                }

                                currentArrayItemIndex = 0;

                                wdbVars.StrArrayDict.Add(wdbVars.NumStringFields[arrayIterator], wdbVars.ProcessStringsList);
                                //Console.WriteLine($"Built [{wdbVars.NumStringFields[arrayIterator]}]");

                                if (arrayIterator + 1 != wdbVars.StrArrayOffsets.Count)
                                {
                                    arrayIterator = wdbVars.StrArrayOffsets.IndexOf((uint)strArrayReader.BaseStream.Position);

                                    buildArray = true;
                                }
                            }
                        }
                    }
                }
            }

            Console.WriteLine($"Finished organizing {wdbVars.StrArraySectionName}");
        }
    }
}