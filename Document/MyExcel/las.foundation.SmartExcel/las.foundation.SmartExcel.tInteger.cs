namespace las.foundation.SmartExcel
{
    using System;
    using System.Runtime.InteropServices;

    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Auto, Pack=1)]
    internal struct tInteger
    {
        public short opcode;
        public short length;
        public short row;
        public short col;
        [MarshalAs(UnmanagedType.I1)]
        public byte rgbAttr1;
        [MarshalAs(UnmanagedType.I1)]
        public byte rgbAttr2;
        [MarshalAs(UnmanagedType.I1)]
        public byte rgbAttr3;
        public short intValue;
    }
}
