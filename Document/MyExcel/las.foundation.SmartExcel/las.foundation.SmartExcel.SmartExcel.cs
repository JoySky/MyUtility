namespace las.foundation.SmartExcel
{
    using System;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Text;

    public class SmartExcel
    {
        private FileStream fs;
        private int[] m_shtHorizPageBreakRows;
        private int m_shtNumHorizPageBreaks = 1;
        private BEG_FILE_RECORD m_udtBEG_FILE_MARKER;
        private END_FILE_RECORD m_udtEND_FILE_MARKER;
        private HPAGE_BREAK_RECORD m_udtHORIZ_PAGE_BREAK;

        public SmartExcel()
        {
            this.Init();
        }

        public void CloseFile()
        {
            try
            {
                if (this.fs != null)
                {
                    if (this.m_shtNumHorizPageBreaks > 0)
                    {
                        for (int i = this.m_shtHorizPageBreakRows.GetUpperBound(0); i >= this.m_shtHorizPageBreakRows.GetLowerBound(0); i--)
                        {
                            for (int k = this.m_shtHorizPageBreakRows.GetLowerBound(0) + 1; k <= i; k++)
                            {
                                if (this.m_shtHorizPageBreakRows[k - 1] > this.m_shtHorizPageBreakRows[k])
                                {
                                    int num4 = this.m_shtHorizPageBreakRows[k - 1];
                                    this.m_shtHorizPageBreakRows[k - 1] = this.m_shtHorizPageBreakRows[k];
                                    this.m_shtHorizPageBreakRows[k] = (short) num4;
                                }
                            }
                        }
                        this.m_udtHORIZ_PAGE_BREAK.opcode = 0x1b;
                        this.m_udtHORIZ_PAGE_BREAK.length = (short) (2 + (this.m_shtNumHorizPageBreaks * 2));
                        this.m_udtHORIZ_PAGE_BREAK.NumPageBreaks = (short) this.m_shtNumHorizPageBreaks;
                        this.FilePut(this.m_udtHORIZ_PAGE_BREAK);
                        for (short j = 1; j <= this.m_shtHorizPageBreakRows.GetUpperBound(0); j = (short) (j + 1))
                        {
                            this.FilePut(Encoding.Default.GetBytes(this.MKI((short) this.m_shtHorizPageBreakRows[j])));
                        }
                    }
                    this.FilePut(this.m_udtEND_FILE_MARKER);
                    this.fs.Close();
                }
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        public void CreateFile(string strFileName)
        {
            try
            {
                if (File.Exists(strFileName))
                {
                    File.SetAttributes(strFileName, FileAttributes.Normal);
                    File.Delete(strFileName);
                }
                this.fs = new FileStream(strFileName, FileMode.CreateNew);
                this.FilePut(this.m_udtBEG_FILE_MARKER);
                this.WriteDefaultFormats();
                this.m_shtHorizPageBreakRows = new int[1];
                this.m_shtNumHorizPageBreaks = 0;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        private void FilePut(byte[] buf)
        {
            this.fs.Write(buf, 0, buf.Length);
        }

        private void FilePut(ValueType vt)
        {
            Type type = vt.GetType();
            int cb = 0;
            cb = Marshal.SizeOf(vt);
            IntPtr ptr = Marshal.AllocHGlobal(cb);
            Marshal.StructureToPtr(vt, ptr, true);
            byte[] destination = new byte[cb];
            Marshal.Copy(ptr, destination, 0, cb);
            this.fs.Write(destination, 0, destination.Length);
            Marshal.FreeHGlobal(ptr);
        }

        private void FilePut(ValueType vt, int len)
        {
            int cb = 0;
            cb = len;
            IntPtr ptr = Marshal.AllocHGlobal(cb);
            Marshal.StructureToPtr(vt, ptr, true);
            byte[] destination = new byte[cb];
            Marshal.Copy(ptr, destination, 0, cb);
            this.fs.Write(destination, 0, destination.Length);
            Marshal.FreeHGlobal(ptr);
        }

        private int GetLength(string strText)
        {
            return Encoding.Default.GetBytes(strText).Length;
        }

        private void Init()
        {
            this.m_udtBEG_FILE_MARKER.opcode = 9;
            this.m_udtBEG_FILE_MARKER.length = 4;
            this.m_udtBEG_FILE_MARKER.version = 2;
            this.m_udtBEG_FILE_MARKER.ftype = 10;
            this.m_udtEND_FILE_MARKER.opcode = 10;
        }

        public void InsertHorizPageBreak(int lrow)
        {
            try
            {
                int num;
                if ((lrow > 0x7fff) || (lrow < 0))
                {
                    num = 0;
                }
                else
                {
                    num = lrow - 1;
                }
                this.m_shtNumHorizPageBreaks++;
                this.m_shtHorizPageBreakRows[this.m_shtNumHorizPageBreaks] = num;
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        private string MKI(short x)
        {
            string lpvDest = "  ";
            RtlMoveMemory(ref lpvDest, ref x, 2);
            return lpvDest;
        }

        [DllImport("kernel32.dll")]
        private static extern void RtlMoveMemory(ref string lpvDest, ref short lpvSource, int cbCopy);
        public void SetColumnWidth(int FirstColumn, int LastColumn, short WidthValue)
        {
            try
            {
                COLWIDTH_RECORD colwidth_record;
                colwidth_record.opcode = 0x24;
                colwidth_record.length = 4;
                colwidth_record.col1 = (byte) (FirstColumn - 1);
                colwidth_record.col2 = (byte) (LastColumn - 1);
                colwidth_record.ColumnWidth = (short) (WidthValue * 0x100);
                this.FilePut(colwidth_record);
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        public void SetDefaultRowHeight(int HeightValue)
        {
            try
            {
                DEF_ROWHEIGHT_RECORD def_rowheight_record;
                def_rowheight_record.opcode = 0x25;
                def_rowheight_record.length = 2;
                def_rowheight_record.RowHeight = HeightValue * 20;
                this.FilePut(def_rowheight_record);
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        public void SetFilePassword(string PasswordText)
        {
            try
            {
                PASSWORD_RECORD password_record;
                int length = this.GetLength(PasswordText);
                password_record.opcode = 0x2f;
                password_record.length = (short) length;
                this.FilePut(password_record);
                this.FilePut(Encoding.Default.GetBytes(PasswordText));
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        public void SetFont(string FontName, short FontHeight, FontFormatting FontFormat)
        {
            try
            {
                FONT_RECORD font_record;
                int length = this.GetLength(FontName);
                font_record.opcode = 0x31;
                font_record.length = (short) (5 + length);
                font_record.FontHeight = (short) (FontHeight * 20);
                font_record.FontAttributes1 = (byte) FontFormat;
                font_record.FontAttributes2 = 0;
                font_record.FontNameLength = (byte) length;
                this.FilePut(font_record);
                this.FilePut(Encoding.Default.GetBytes(FontName));
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        public void SetFooter(string FooterText)
        {
            try
            {
                HEADER_FOOTER_RECORD header_footer_record;
                int length = this.GetLength(FooterText);
                header_footer_record.opcode = 0x15;
                header_footer_record.length = (short) (1 + length);
                header_footer_record.TextLength = (byte) length;
                this.FilePut(header_footer_record);
                this.FilePut(Encoding.Default.GetBytes(FooterText));
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        public void SetHeader(string HeaderText)
        {
            try
            {
                HEADER_FOOTER_RECORD header_footer_record;
                int length = this.GetLength(HeaderText);
                header_footer_record.opcode = 20;
                header_footer_record.length = (short) (1 + length);
                header_footer_record.TextLength = (byte) length;
                this.FilePut(header_footer_record);
                this.FilePut(Encoding.Default.GetBytes(HeaderText));
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        public void SetMargin(MarginTypes Margin, double MarginValue)
        {
            try
            {
                MARGIN_RECORD_LAYOUT margin_record_layout;
                margin_record_layout.opcode = (short) Margin;
                margin_record_layout.length = 8;
                margin_record_layout.MarginValue = MarginValue;
                this.FilePut(margin_record_layout);
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        public void SetRowHeight(int Row, short HeightValue)
        {
            if (Row > 0x7fff)
            {
                throw new Exception("行高不能大于32767");
            }
            try
            {
                ROW_HEIGHT_RECORD row_height_record;
                int num = Row;
                row_height_record.opcode = 8;
                row_height_record.length = 0x10;
                row_height_record.RowNumber = num;
                row_height_record.FirstColumn = 0;
                row_height_record.LastColumn = 0x100;
                row_height_record.RowHeight = HeightValue * 20;
                row_height_record.internals = 0;
                row_height_record.DefaultAttributes = 0;
                row_height_record.FileOffset = 0;
                row_height_record.rgbAttr1 = 0;
                row_height_record.rgbAttr2 = 0;
                row_height_record.rgbAttr3 = 0;
                this.FilePut(row_height_record);
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        private void WriteDefaultFormats()
        {
            FORMAT_COUNT_RECORD format_count_record;
            string str = "\"";
            string[] strArray = new string[] { 
                "General", "0", "0.00", "#,##0", "#,##0.00", @"#,##0\ " + str + "$" + str + @";\-#,##0\ " + str + "$" + str, @"#,##0\ " + str + "$" + str + @";[Red]\-#,##0\ " + str + "$" + str, @"#,##0.00\ " + str + "$" + str + @";\-#,##0.00\ " + str + "$" + str, @"#,##0.00\ " + str + "$" + str + @";[Red]\-#,##0.00\ " + str + "$" + str, "0%", "0.00%", "0.00E+00", "dd/mm/yy", @"dd/\ mmm\ yy", @"dd/\ mmm", @"mmm\ yy", 
                @"h:mm\ AM/PM", @"h:mm:ss\ AM/PM", "hh:mm", "hh:mm:ss", @"dd/mm/yy\ hh:mm", "##0.0E+0", "mm:ss", "@"
             };
            format_count_record.opcode = 0x1f;
            format_count_record.length = 2;
            format_count_record.Count = (short) strArray.GetUpperBound(0);
            this.FilePut(format_count_record);
            for (int i = strArray.GetLowerBound(0); i <= strArray.GetUpperBound(0); i++)
            {
                FORMAT_RECORD format_record;
                int length = strArray[i].Length;
                format_record.opcode = 30;
                format_record.length = (short) (length + 1);
                format_record.FormatLength = (byte) length;
                this.FilePut(format_record);
                for (int j = 0; j < length; j++)
                {
                    byte num3 = (byte) strArray[i].Substring(j, 1).ToCharArray(0, 1)[0];
                    this.FilePut(new byte[] { num3 });
                }
            }
        }

        public void WriteValue(ValueTypes ValueType, CellFont CellFontUsed, CellAlignment Alignment, CellHiddenLocked HiddenLocked, int lrow, int lcol, object Value)
        {
            this.WriteValue(ValueType, CellFontUsed, Alignment, HiddenLocked, lrow, lcol, Value, 0);
        }

        public void WriteValue(ValueTypes ValueType, CellFont CellFontUsed, CellAlignment Alignment, CellHiddenLocked HiddenLocked, int lrow, int lcol, object Value, int CellFormat)
        {
            try
            {
                short num2;
                short num3;
                if ((lrow > 0x7fff) || (lrow < 0))
                {
                    num3 = 0;
                }
                else
                {
                    num3 = (short) (lrow - 1);
                }
                if ((lcol > 0x7fff) || (lcol < 0))
                {
                    num2 = 0;
                }
                else
                {
                    num2 = (short) (lcol - 1);
                }
                switch (ValueType)
                {
                    case ValueTypes.Integer:
                        tInteger integer;
                        integer.opcode = 2;
                        integer.length = 9;
                        integer.row = num3;
                        integer.col = num2;
                        integer.rgbAttr1 = (byte) HiddenLocked;
                        integer.rgbAttr2 = (byte) (CellFontUsed + CellFormat);
                        integer.rgbAttr3 = (byte) Alignment;
                        integer.intValue = (short) Value;
                        this.FilePut(integer);
                        return;

                    case ValueTypes.Number:
                        tNumber number;
                        number.opcode = 3;
                        number.length = 15;
                        number.row = num3;
                        number.col = num2;
                        number.rgbAttr1 = (byte) HiddenLocked;
                        number.rgbAttr2 = (byte) (CellFontUsed + CellFormat);
                        number.rgbAttr3 = (byte) Alignment;
                        number.NumberValue = (double) Value;
                        this.FilePut(number);
                        return;

                    case ValueTypes.Text:
                    {
                        tText text;
                        string strText = Convert.ToString(Value);
                        int length = this.GetLength(strText);
                        text.opcode = 4;
                        text.length = 10;
                        text.TextLength = (byte) length;
                        text.length = (byte) (8 + length);
                        text.row = num3;
                        text.col = num2;
                        text.rgbAttr1 = (byte) HiddenLocked;
                        text.rgbAttr2 = (byte) (CellFontUsed + CellFormat);
                        text.rgbAttr3 = (byte) Alignment;
                        this.FilePut(text);
                        this.FilePut(Encoding.Default.GetBytes(strText));
                        return;
                    }
                }
            }
            catch (Exception exception)
            {
                throw exception;
            }
        }

        public bool PrintGridLines
        {
            set
            {
                try
                {
                    PRINT_GRIDLINES_RECORD print_gridlines_record;
                    print_gridlines_record.opcode = 0x2b;
                    print_gridlines_record.length = 2;
                    if (value)
                    {
                        print_gridlines_record.PrintFlag = 1;
                    }
                    else
                    {
                        print_gridlines_record.PrintFlag = 0;
                    }
                    this.FilePut(print_gridlines_record);
                }
                catch (Exception exception)
                {
                    throw exception;
                }
            }
        }

        public bool ProtectSpreadsheet
        {
            set
            {
                try
                {
                    PROTECT_SPREADSHEET_RECORD protect_spreadsheet_record;
                    protect_spreadsheet_record.opcode = 0x12;
                    protect_spreadsheet_record.length = 2;
                    if (value)
                    {
                        protect_spreadsheet_record.Protect = 1;
                    }
                    else
                    {
                        protect_spreadsheet_record.Protect = 0;
                    }
                    if (this.fs == null)
                    {
                        throw new SmartExcelOpeartionFileException();
                    }
                    this.FilePut(protect_spreadsheet_record);
                }
                catch (Exception exception)
                {
                    throw exception;
                }
            }
        }
    }
}
