namespace las.foundation.SmartExcel
{
    using System;
    using System.Runtime.InteropServices;

    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Auto, Pack=1)]
    internal struct MARGIN_RECORD_LAYOUT
    {
        public short opcode;
        public short length;
        public double MarginValue;
    }
}
