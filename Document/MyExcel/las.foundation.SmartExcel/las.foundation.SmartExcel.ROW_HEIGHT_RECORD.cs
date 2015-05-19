namespace las.foundation.SmartExcel
{
    using System;
    using System.Runtime.InteropServices;

    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Auto, Pack=1)]
    internal struct ROW_HEIGHT_RECORD
    {
        public int opcode;
        public int length;
        public int RowNumber;
        public int FirstColumn;
        public int LastColumn;
        public int RowHeight;
        public int internals;
        [MarshalAs(UnmanagedType.I1)]
        public byte DefaultAttributes;
        public int FileOffset;
        [MarshalAs(UnmanagedType.I1)]
        public byte rgbAttr1;
        [MarshalAs(UnmanagedType.I1)]
        public byte rgbAttr2;
        [MarshalAs(UnmanagedType.I1)]
        public byte rgbAttr3;
    }
}
