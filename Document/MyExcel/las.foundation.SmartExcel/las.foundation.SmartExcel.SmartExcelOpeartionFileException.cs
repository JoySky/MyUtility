namespace las.foundation.SmartExcel
{
    using System;

    public class SmartExcelOpeartionFileException : ApplicationException
    {
        public SmartExcelOpeartionFileException() : base("请首先调用CreateFile方法!")
        {
        }
    }
}
