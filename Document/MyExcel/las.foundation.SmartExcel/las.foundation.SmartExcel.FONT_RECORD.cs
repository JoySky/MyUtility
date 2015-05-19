namespace las.foundation.SmartExcel
{
    using System;
    using System.Runtime.InteropServices;

    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Auto, Pack=1)]
    internal struct FONT_RECORD
    {
        public short opcode;
        public short length;
        public short FontHeight;
        [MarshalAs(UnmanagedType.I1)]
        public byte FontAttributes1;
        [MarshalAs(UnmanagedType.I1)]
        public byte FontAttributes2;
        [MarshalAs(UnmanagedType.I1)]
        public byte FontNameLength;
    }
}
