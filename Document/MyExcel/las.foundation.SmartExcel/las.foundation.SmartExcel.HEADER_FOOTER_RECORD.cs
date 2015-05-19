namespace las.foundation.SmartExcel
{
    using System;
    using System.Runtime.InteropServices;

    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Auto, Pack=1)]
    internal struct HEADER_FOOTER_RECORD
    {
        public short opcode;
        public short length;
        [MarshalAs(UnmanagedType.I1)]
        public byte TextLength;
    }
}
