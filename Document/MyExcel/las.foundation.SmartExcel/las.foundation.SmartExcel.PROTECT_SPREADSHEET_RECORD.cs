namespace las.foundation.SmartExcel
{
    using System;
    using System.Runtime.InteropServices;

    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Auto, Pack=1)]
    internal struct PROTECT_SPREADSHEET_RECORD
    {
        public short opcode;
        public short length;
        public short Protect;
    }
}
