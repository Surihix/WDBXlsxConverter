using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace WDBXlsxConverter
{
    internal class WDBMethods
    {
        public static void ErrorExit(string errorMsg)
        {
            Console.WriteLine("");
            Console.WriteLine($"Error: {errorMsg}");
            Console.ReadLine();
            Environment.Exit(1);
        }


        public static void CheckSectionName(BinaryReader br, string sectionName)
        {
            if (br.ReadBytesString(16, false) != sectionName)
            {
                ErrorExit($"{sectionName} is not present in the expected position");
            }
        }


        public static byte[] SaveSectionData(BinaryReader br, bool reverse)
        {
            var sectionOffset = br.ReadBytesUInt32(true);
            var sectionLength = br.ReadBytesUInt32(true);

            _ = br.BaseStream.Position = sectionOffset;
            var sectionData = br.ReadBytes((int)sectionLength);

            if (reverse)
            {
                Array.Reverse(sectionData);
            }

            return sectionData;
        }


        public static List<uint> GetSectionDataValues(byte[] dataArray)
        {
            var processList = new List<uint>();
            var dataIndex = 0;

            for (int i = 0; i < dataArray.Length / 4; i++)
            {
                var currentValue = DeriveUIntFromSectionData(dataArray, dataIndex, true);
                processList.Add(currentValue);

                dataIndex += 4;
            }

            return processList;
        }


        public static void WriteValuesListToSheet(List<uint> valuesList, IXLWorksheet worksheet, int cellX, int cellY)
        {
            for (int i = 0; i < valuesList.Count; i++)
            {
                WriteToSheet(worksheet, cellX, cellY, valuesList[i].ToString(), 3, false);
                cellX++;
            }
        }


        public static string DeriveStringFromArray(byte[] dataArray, int stringOffset)
        {
            var length = 0;
            for (int s = stringOffset; s < dataArray.Length; s++)
            {
                if (dataArray[s] == 0)
                {
                    break;
                }

                length++;
            }

            return Encoding.UTF8.GetString(dataArray, stringOffset, length);
        }


        public static int DeriveFieldNumber(string fieldName)
        {
            var foundNumsList = new List<int>();

            for (int i = 1; i < 3; i++)
            {
                if (i == 1 && !char.IsDigit(fieldName[i]))
                {
                    break;
                }

                if (char.IsDigit(fieldName[i]))
                {
                    foundNumsList.Add(int.Parse(Convert.ToString(fieldName[i])));
                }
            }

            var foundNumStr = "";
            foreach (var n in foundNumsList)
            {
                foundNumStr += n;
            }

            var hasParsed = int.TryParse(foundNumStr, out int foundNum);

            if (hasParsed)
            {
                return foundNum;
            }
            else
            {
                return 0;
            }
        }


        public static void WriteToSheet(IXLWorksheet workSheet, int cellX, int cellY, object valueToWrite, int typeValue, bool bold)
        {
            switch (typeValue)
            {
                // float value
                case 1:
                    workSheet.Cell(cellX, cellY).Value = Convert.ToDecimal(valueToWrite);
                    workSheet.Cell(cellX, cellY).Style.Font.Bold = bold;
                    workSheet.Cell(cellX, cellY).Style.Font.FontName = "Arial";
                    break;

                // bitpacked and
                // string values
                case 2:
                    workSheet.Cell(cellX, cellY).Value = valueToWrite.ToString();
                    workSheet.Cell(cellX, cellY).Style.Font.Bold = bold;
                    workSheet.Cell(cellX, cellY).Style.Font.FontName = "Arial";
                    break;

                // uint value
                case 3:
                    workSheet.Cell(cellX, cellY).Value = Convert.ToUInt32(valueToWrite);
                    workSheet.Cell(cellX, cellY).Style.Font.Bold = bold;
                    workSheet.Cell(cellX, cellY).Style.Font.FontName = "Arial";
                    break;

                // int value
                case 4:
                    workSheet.Cell(cellX, cellY).Value = Convert.ToInt32(valueToWrite);
                    workSheet.Cell(cellX, cellY).Style.Font.Bold = bold;
                    workSheet.Cell(cellX, cellY).Style.Font.FontName = "Arial";
                    break;

                // uint64
                case 5:
                    workSheet.Cell(cellX, cellY).Value = Convert.ToUInt64(valueToWrite);
                    workSheet.Cell(cellX, cellY).Style.Font.Bold = bold;
                    workSheet.Cell(cellX, cellY).Style.Font.FontName = "Arial";
                    break;
            }
        }


        public static uint DeriveUIntFromSectionData(byte[] dataArray, int dataArrayIndex, bool reverse)
        {
            var processArray = new byte[4];
            Array.ConstrainedCopy(dataArray, dataArrayIndex, processArray, 0, 4);

            if (reverse)
            {
                Array.Reverse(processArray);
            }

            return BitConverter.ToUInt32(processArray, 0);
        }


        public static float DeriveFloatFromSectionData(byte[] dataArray, int dataArrayIndex, bool reverse)
        {
            var processArray = new byte[4];
            Array.ConstrainedCopy(dataArray, dataArrayIndex, processArray, 0, 4);

            if (reverse)
            {
                Array.Reverse(processArray);
            }

            return BitConverter.ToSingle(processArray, 0);
        }
    }
}