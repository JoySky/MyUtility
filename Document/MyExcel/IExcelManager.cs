/* -----------------------------------------------------------------------
 *    ��Ȩ���У� ��Ȩ����(C) 2006��EYE��Ver1.1��
 *    �ļ���ţ� 0003
 *    �ļ����ƣ� IExcelManager.cs
 *    ϵͳ��ţ� E-Eye_0001
 *    ϵͳ���ƣ� SFWL����ϵͳ
 *    ģ���ţ� 0001
 *    ģ�����ƣ� Excel���ݹ������ӿ�
 *    ����ĵ��� 
 *    ������ڣ� 2006-5-12 10:33:56
 *    �������ߣ� �ָ���
 *    ����ժҪ�� ���Excel���ݹ������API�ӿڶ���
 *    ���������� �ýӿ���10������
 *					01 DataSource			����Դ
 *					02 Title				Excel�ļ�������
 *					03 FilePath				Դ�ļ�·��
 *					04 XMLFilePath			XML�ܹ��ļ�·��
 *	  			    05 IsOpen				�ж�Excel��������Ƿ��Ѿ���
 *					06 WriteType			д������(��ͨ����д��׷��)
 *					07 SheetName			Sheet������
 *					08 BackColor			������ɫ
 *					09 ForeColor			������ɫ
 *					10 Font					������ʽ
 *    ���������� �ýӿڰ���10������:
 *					01 Open					��Excel�������
 *					02 Read					��ȡExcel�ļ�����Ч�������ݼ�
 *					03 ReadCell				��ȡĳ��Ԫ�������
 *					04 ReadCell				��ȡĳ��Ԫ������ݣ����У��з�ʽ
 *					05 Write				д����Excel�ļ���
 *					06 ReWrite				����ָ����д����д����
 *					07 WriteCell			д��������Ԫ��
 *					08 Close				��Դ�ͷ�
 *					09 OpenCreate			��Excel���������Excel�ļ��������򴴽�֮
 *					10 ActiveSheet			���ǰ��д���ݵ�Sheet��
 *    �ļ����ã���
 *	  Լ    ����1.ִ�е�����������Ҫװ��Office�������Excel�ļ��汾��2000����
 *				2.Excel���ļ�Ŀǰ��֧�ֵ������������������ȡ
 *				3.���ȡExcel�ļ����ڵ�һ�б������δ洢������Χ�������޶�����״̬λ������Ϊ����Ҫ��ȡ����ʼ��Ԫ�����ƣ�������Ԫ������
 *				4.���ȡ��ÿ��Excel�����ļ�����������֮��Ե�ͬ��XML�ܹ��ļ�����չ���֧�ֲ�ͬ�ļ��������Ƽ����
 *				5.д�ļ�ʱ���Ƽ��Ƚ�����Excel�ļ�����չ���֧���Զ�����Excel�ļ������Ƽ����,���汾�ݲ��ṩ���ݲ��빦��
 *				6.д����Excel�ļ����ṩ����ͨ��Insert��д��Append׷�ӣ����ֲ�����ʽ
 *					�������Ĵ�СҪ�󵥴�д�룺��1--60000����,��1-255��Cellֵ����1-255�ַ� 
 *					Ŀǰд������֧�ֵ���������������32��Sheet��ÿSheet���洢��Ϊ60000��
 *				7.����Լ������.Net��ܼ�Microsoft Office Excel���Լ����
 *    -----------------------------------------------------------------------
 * */

using System;
using System.Collections;
using System.Data;

namespace MyUtility.OFFICE.MyExcel
{
	/// �� �� �ţ� 01
	/// �� �� �ƣ� IExcelManager
	/// ����ժҪ�� Excel�������ӿ�
	/// ������ڣ� 2006-5-12 14:36:51
	/// �������ߣ� �ָ���
	public interface IExcelManager : IDisposable
	{
		#region Property
		/// <summary>
		/// ����Դ
		/// </summary>
		DataSet DataSource{get;set;}
		/// <summary>
		/// Excel�ļ�������
		/// </summary>
		string Title{get;set;}
		/// <summary>
		/// Դ�ļ�·��
		/// </summary>
		string FilePath{get;set;}
		/// <summary>
		/// XML�ܹ��ļ�·��
		/// </summary>
		string XMLFilePath{get;set;}
		/// <summary>
		/// �ж�Excel��������Ƿ��Ѿ���
		/// </summary>
		/// <returns></returns>
		bool IsOpen{get;}
		/// <summary>
		/// Excel�ļ�д������
		/// </summary>
		EnumType.WriteType WriteType{get;set;}
		/// <summary>
		/// Sheet������
		/// </summary>
		string SheetName{get;set;}
		/// <summary>
		/// ������ɫ
		/// </summary>
		System.Drawing.Color BackColor{get;set;}
		/// <summary>
		/// ������ɫ
		/// </summary>
		System.Drawing.Color ForeColor{get;set;}
		/// <summary>
		/// ������ʽ
		/// </summary>
		System.Drawing.Font Font{get;set;}
		#endregion

		#region Private Method
		/// <summary>
		/// ��Excel�������
		/// </summary>
		/// <returns></returns>
		bool Open();
		/// <summary>
		/// ��ȡExcel�ļ�����Ч�������ݼ�
		/// </summary>
		/// <returns></returns>
		DataSet Read();
		/// <summary>
		/// ��ȡĳ��Ԫ�������
		/// </summary>
		/// <param name="strCell">��Ԫ������</param>
		/// <returns></returns>
		string ReadCell(string strCell);
		/// <summary>
		/// ��ȡĳ��Ԫ�����ݣ������в�����ʽ
		/// </summary>
		/// <param name="iRow"></param>
		/// <param name="iCol"></param>
		/// <returns></returns>
		string ReadCell(int iRow,int iCol);
		/// <summary>
		/// д�����ݼ���Excel�ļ�
		/// </summary>
		/// <returns></returns>
		bool Write();
		/// <summary>
		/// ��д���ݼ���Excel�ļ�,ָ����д��
		/// </summary>
		/// <param name="iRow"></param>
		/// <returns></returns>
		bool ReWrite(int iRow);
		/// <summary>
		/// д��������Ԫ��
		/// </summary>
		bool WriteCell(int iRow,int iCol,string strValue);
		/// <summary>
		/// ���Sheet��
		/// </summary>
		/// <param name="sheetindex"></param>
		void ActiveSheet(EnumType.SheetIndex sheetindex);
		/// <summary>
		/// ��Դ�ͷ�
		/// </summary>
		void Close();
		/// <summary>
		/// �򿪲������˶���
		/// </summary>
		/// <returns></returns>
		bool OpenCreate();
		#endregion
	}
}
