using System;

namespace MyUtility.OFFICE.MyExcel
{
	/// <summary>
	/// 文件写入操作类型
	/// </summary>
	public class EnumType
	{
		public enum WriteType
		{
			/// <summary>
			/// 默认类型，写入新文件
			/// </summary>
			None,
			/// <summary>
			/// 重写
			/// </summary>
			ReWrite,
			/// <summary>
			/// 追加
			/// </summary>
			Append,
			/// <summary>
			/// 插入
			/// </summary>
			Insert
		}
		/// <summary>
		/// Sheet表定位,最大32个Sheet
		/// </summary>
		public enum SheetIndex
		{
			Sheet1 = 1,
			Sheet2 = 2,
			Sheet3 = 3,
			Sheet4 = 4,
			Sheet5 = 5,
			Sheet6 = 6,
			Sheet7 = 7,
			Sheet8 = 8,
			Sheet9 = 9,
			Sheet10 = 10,
			Sheet11 = 11,
			Sheet12 = 12,
			Sheet13 = 13,
			Sheet14 = 14,
			Sheet15 = 15,
			Sheet16 = 16,
			Sheet17 = 17,
			Sheet18 = 18,
			Sheet19 = 19, 
			Sheet20 = 20,
			Sheet21 = 21,
			Sheet22 = 22,
			Sheet23 = 23,
			Sheet24 = 24,
			Sheet25 = 25,
			Sheet26 = 26,
			Sheet27 = 27,
			Sheet28 = 28,
			Sheet29 = 29,
			Sheet30 = 30,
			Sheet31 = 31,
			Sheet32 = 32,
		}
	}


}
