/* -----------------------------------------------------------------------
 *    版权所有： 版权所有(C) 2006，EYE
 *    文件编号： 0004
 *    文件名称： ExcelManagerFactory.cs
 *    系统编号： E-Eye_0001
 *    系统名称： SFWL管理系统
 *    模块编号： 0001
 *    模块名称： Excel数据管理对象工厂类
 *    设计文档： 
 *    完成日期： 2006-5-12 10:33:56
 *    作　　者： 林付国
 *    内容摘要： 完成Excel数据管理对象创建
 *    属性描述： 无
 *			  
 *    方法描述： 该类有2个方法:
 *					01 Instance					返回对象实例
 *					02 CreateExcelManager		创建Excel管理对象
 *    文件调用：无
 *    -----------------------------------------------------------------------
 * */

using System;

namespace MyUtility.OFFICE.MyExcel
{
	/// <summary>
	///  类 编 号： 03
	///  类 名 称： ExcelManagerFactory 
	///  内容摘要： 完成Excel数据管理对象创建
	///  完成日期： 2006-5-12 14:38:38
	///  编码作者： 林付国
	/// </summary>
	public class ExcelManagerFactory
	{
		#region 成员变量

		private static ExcelManagerFactory m_ExcelManagerFactory = new ExcelManagerFactory();

		#endregion

		#region Private Method

		private ExcelManagerFactory(){}

		/// <summary>
		/// 返回对象实例
		/// </summary>
		/// <returns></returns>
		public static ExcelManagerFactory Instance()
		{
			return m_ExcelManagerFactory;
		}
		/// <summary>
		/// 创建Excel管理对象
		/// </summary>
		/// <param name="strpath">源文件路径</param>
		/// <returns></returns>
		public IExcelManager CreateExcelManager()
		{
			return p_CreateExcelManager();
		}

		private IExcelManager p_CreateExcelManager()
		{
            return new CExcelManager();
		}
		
		#endregion
	}
}
