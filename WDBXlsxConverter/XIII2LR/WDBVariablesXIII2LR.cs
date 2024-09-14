﻿using System.Collections.Generic;

namespace WDBXlsxConverter.XIII2LR
{
    internal class WDBVariablesXIII2LR
    {
        // Important variables
        public string WDBName;
        public string xlsxName;
        public uint RecordCount;
        public bool HasStrArraySection;
        public bool HasStringSection;
        public bool parseStrtypelistAsV1;
        public bool hasTypelistSection;
        public string[] Fields;
        public List<uint> StrArrayOffsets = new List<uint>();
        public List<string> NumStringFields = new List<string>();
        public List<string> ProcessStringsList = new List<string>();
        public Dictionary<string, List<string>> StrArrayDict = new Dictionary<string, List<string>>();
        public List<int> StrtypelistValues = new List<int>();

        // Section names
        public string SheetNameSectionName = "!!sheetname";
        public string StrArraySectionName = "!!strArray";
        public string StrArrayInfoSectionName = "!!strArrayInfo";
        public string StrArrayListSectionName = "!!strArrayList";
        public string StringSectionName = "!!string";
        public string StrtypelistSectionName = "!!strtypelist";
        public string StrtypelistbSectionName = "!!strtypelistb";
        public string TypelistSectionName = "!!typelist";
        public string VersionSectionName = "!!version";
        public string StructItemSectionName = "!structitem";
        public string StructItemNumSectionName = "!structitemnum";

        // Section data
        public string SheetName;
        public byte[] StrArrayData;
        public byte OffsetsPerValue;
        public byte BitsPerOffset;
        public byte[] StrArrayListData;
        public byte[] StringsData;
        public byte[] StrtypelistData;
        public byte[] TypelistData;
        public byte[] VersionData;
        public byte[] StructItemData;
        public uint FieldCount;
    }
}