namespace las.foundation.SmartExcel
{
    using System;
    using System.Runtime.InteropServices;

    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Auto, Pack=1)]
    internal struct DEF_ROWHEIGHT_RECORD
    {
        public int opcode;
        public int length;
        public int RowHeight;
    }
}
