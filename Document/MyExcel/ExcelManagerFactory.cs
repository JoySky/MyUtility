/* -----------------------------------------------------------------------
 *    ��Ȩ���У� ��Ȩ����(C) 2006��EYE
 *    �ļ���ţ� 0004
 *    �ļ����ƣ� ExcelManagerFactory.cs
 *    ϵͳ��ţ� E-Eye_0001
 *    ϵͳ���ƣ� SFWL����ϵͳ
 *    ģ���ţ� 0001
 *    ģ�����ƣ� Excel���ݹ�����󹤳���
 *    ����ĵ��� 
 *    ������ڣ� 2006-5-12 10:33:56
 *    �������ߣ� �ָ���
 *    ����ժҪ�� ���Excel���ݹ�����󴴽�
 *    ���������� ��
 *			  
 *    ���������� ������2������:
 *					01 Instance					���ض���ʵ��
 *					02 CreateExcelManager		����Excel�������
 *    �ļ����ã���
 *    -----------------------------------------------------------------------
 * */

using System;

namespace MyUtility.OFFICE.MyExcel
{
	/// <summary>
	///  �� �� �ţ� 03
	///  �� �� �ƣ� ExcelManagerFactory 
	///  ����ժҪ�� ���Excel���ݹ�����󴴽�
	///  ������ڣ� 2006-5-12 14:38:38
	///  �������ߣ� �ָ���
	/// </summary>
	public class ExcelManagerFactory
	{
		#region ��Ա����

		private static ExcelManagerFactory m_ExcelManagerFactory = new ExcelManagerFactory();

		#endregion

		#region Private Method

		private ExcelManagerFactory(){}

		/// <summary>
		/// ���ض���ʵ��
		/// </summary>
		/// <returns></returns>
		public static ExcelManagerFactory Instance()
		{
			return m_ExcelManagerFactory;
		}
		/// <summary>
		/// ����Excel�������
		/// </summary>
		/// <param name="strpath">Դ�ļ�·��</param>
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
