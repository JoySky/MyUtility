/* -----------------------------------------------------------------------
 *    版权所有： 版权所有(C) 2006，EYE（Ver1.1）
 *    文件编号： 0003
 *    文件名称： IExcelManager.cs
 *    系统编号： E-Eye_0001
 *    系统名称： SFWL管理系统
 *    模块编号： 0001
 *    模块名称： Excel数据管理对象接口
 *    设计文档： 
 *    完成日期： 2006-5-12 10:33:56
 *    作　　者： 林付国
 *    内容摘要： 完成Excel数据管理对象API接口定义
 *    属性描述： 该接口有10个属性
 *					01 DataSource			数据源
 *					02 Title				Excel文件表格标题
 *					03 FilePath				源文件路径
 *					04 XMLFilePath			XML架构文件路径
 *	  			    05 IsOpen				判断Excel管理对象是否已经打开
 *					06 WriteType			写入类型(普通，重写，追加)
 *					07 SheetName			Sheet表名称
 *					08 BackColor			背景颜色
 *					09 ForeColor			字体颜色
 *					10 Font					字体样式
 *    方法描述： 该接口包括10个方法:
 *					01 Open					打开Excel管理对象
 *					02 Read					读取Excel文件中有效行列数据集
 *					03 ReadCell				读取某单元格的内容
 *					04 ReadCell				读取某单元格的内容，按行，列方式
 *					05 Write				写入至Excel文件中
 *					06 ReWrite				按照指定重写行重写数据
 *					07 WriteCell			写数据至单元格
 *					08 Close				资源释放
 *					09 OpenCreate			打开Excel管理对象，若Excel文件不存在则创建之
 *					10 ActiveSheet			激活当前读写数据的Sheet表
 *    文件调用：无
 *	  约    束：1.执行导出机器上需要装有Office组件，且Excel文件版本在2000以上
 *				2.Excel读文件目前仅支持单工作簿，单工作表读取
 *				3.需读取Excel文件，在第一行必须依次存储二个范围，用于限定参数状态位，依次为：需要读取的起始单元格名称，结束单元格名称
 *				4.需读取的每个Excel数据文件，必须有于之配对的同名XML架构文件（扩展项可支持不同文件名，不推荐此项）
 *				5.写文件时，推荐先建立空Excel文件（扩展项可支持自动创建Excel文件，不推荐此项）,本版本暂不提供数据插入功能
 *				6.写入至Excel文件，提供（普通，Insert重写，Append追加）三种操作方式
 *					数据量的大小要求单次写入：行1--60000以内,列1-255，Cell值长度1-255字符 
 *					目前写操作仅支持单工作簿，最大存在32个Sheet，每Sheet最大存储量为60000行
 *				7.其它约束按照.Net框架及Microsoft Office Excel相关约定。
 *    -----------------------------------------------------------------------
 * */

using System;
using System.Collections;
using System.Data;

namespace MyUtility.OFFICE.MyExcel
{
	/// 类 编 号： 01
	/// 类 名 称： IExcelManager
	/// 内容摘要： Excel管理对象接口
	/// 完成日期： 2006-5-12 14:36:51
	/// 编码作者： 林付国
	public interface IExcelManager : IDisposable
	{
		#region Property
		/// <summary>
		/// 数据源
		/// </summary>
		DataSet DataSource{get;set;}
		/// <summary>
		/// Excel文件表格标题
		/// </summary>
		string Title{get;set;}
		/// <summary>
		/// 源文件路径
		/// </summary>
		string FilePath{get;set;}
		/// <summary>
		/// XML架构文件路径
		/// </summary>
		string XMLFilePath{get;set;}
		/// <summary>
		/// 判断Excel管理对象是否已经打开
		/// </summary>
		/// <returns></returns>
		bool IsOpen{get;}
		/// <summary>
		/// Excel文件写入类型
		/// </summary>
		EnumType.WriteType WriteType{get;set;}
		/// <summary>
		/// Sheet表名称
		/// </summary>
		string SheetName{get;set;}
		/// <summary>
		/// 背景颜色
		/// </summary>
		System.Drawing.Color BackColor{get;set;}
		/// <summary>
		/// 字体颜色
		/// </summary>
		System.Drawing.Color ForeColor{get;set;}
		/// <summary>
		/// 字体样式
		/// </summary>
		System.Drawing.Font Font{get;set;}
		#endregion

		#region Private Method
		/// <summary>
		/// 打开Excel管理对象
		/// </summary>
		/// <returns></returns>
		bool Open();
		/// <summary>
		/// 读取Excel文件中有效行列数据集
		/// </summary>
		/// <returns></returns>
		DataSet Read();
		/// <summary>
		/// 读取某单元格的内容
		/// </summary>
		/// <param name="strCell">单元格名称</param>
		/// <returns></returns>
		string ReadCell(string strCell);
		/// <summary>
		/// 读取某单元格内容，按行列参数方式
		/// </summary>
		/// <param name="iRow"></param>
		/// <param name="iCol"></param>
		/// <returns></returns>
		string ReadCell(int iRow,int iCol);
		/// <summary>
		/// 写入数据集到Excel文件
		/// </summary>
		/// <returns></returns>
		bool Write();
		/// <summary>
		/// 重写数据集到Excel文件,指定重写行
		/// </summary>
		/// <param name="iRow"></param>
		/// <returns></returns>
		bool ReWrite(int iRow);
		/// <summary>
		/// 写数据至单元格
		/// </summary>
		bool WriteCell(int iRow,int iCol,string strValue);
		/// <summary>
		/// 活动的Sheet表
		/// </summary>
		/// <param name="sheetindex"></param>
		void ActiveSheet(EnumType.SheetIndex sheetindex);
		/// <summary>
		/// 资源释放
		/// </summary>
		void Close();
		/// <summary>
		/// 打开并创建此对象
		/// </summary>
		/// <returns></returns>
		bool OpenCreate();
		#endregion
	}
}
