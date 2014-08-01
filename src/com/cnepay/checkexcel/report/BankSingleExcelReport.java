package com.cnepay.checkexcel.report;

import java.io.File;

/**
 * 银行单个账户报表
 * - 银行上报的单账户表格文件名为：B_支付机构名_银行账号_期数.xls
 * - 银行单个账户表格可能有多张，每张表格包括2个sheet，分别为2-1、2-2，其中2-2暂不使用
 */
public class BankSingleExcelReport extends BaseExcelReport {

	public BankSingleExcelReport(File file) {
		super(file);

		// 命名以B开头
		this.fileType = BaseExcelReport.TYPE_B;
		// 命名使用下划线分为4节
		this.fileNameSectionsNumber = 4;

		this.reportName = "银行单个账户报表";
		this.sheetNames = new String[2];
		sheetNames[0] = "2-1";
		sheetNames[1] = "2-2";
	}
	
}
