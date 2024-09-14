using System.Collections.Generic;

namespace WDBXlsxConverter.XIII
{
    internal class WDBVariablesXIII
    {
        // Important variables
        public string xlsxName;
        public string WDBName;
        public uint RecordCount;
        public bool HasStringSection;
        public List<uint> StrtypelistValues = new List<uint>();
        public List<uint> TypelistValues = new List<uint>();
        public bool IsKnown;
        public bool IgnoreKnown;
        public string SheetName;
        public string[] Fields;

        // Section names
        public string StringSectionName = "!!string";
        public string StrtypelistSectionName = "!!strtypelist";
        public string TypelistSectionName = "!!typelist";
        public string VersionSectionName = "!!version";

        // Section data
        public byte[] StringsData;
        public byte[] StrtypelistData;
        public byte[] TypelistData;
        public byte[] VersionData;
        public uint FieldCount;
    }
}