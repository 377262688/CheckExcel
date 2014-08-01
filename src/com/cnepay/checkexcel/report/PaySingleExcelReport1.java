package com.cnepay.checkexcel.report;

import java.io.File;

/**
 * 支付机构单个账户报表
 * - 文件名为：A_支付机构名_银行账号_期数.xls
 * - 支付机构单个账户表格可能有多张，每张表格包括5个sheet，分别为1-1、1-3、1-6、1-9和1-10
 * - 逐表格逐项核对1-1中的合计数是否为本月所有日期的该项数据之和。逐项核对1-2、1-4中的合计数是否为本月所有日期的该项数据之和
 * 
 */
public class PaySingleExcelReport1 extends BaseReport {
	
	public PaySingleExcelReport1(File file) {
		super(file);
		
		// 命名以A开头
		this.fileType = BaseReport.TYPE_A;
		// 命名使用下划线分为4节
		this.fileNameSectionsNumber = 4;		
		// 报表类型
		this.reportName = "支付机构单个账户报表";
		
		// Sheet名称
		this.sheetNames = new String[5];
		sheetNames[0] = "1-1";
		sheetNames[1] = "1-3";
		sheetNames[2] = "1-6";
		sheetNames[3] = "1-9";
		sheetNames[4] = "1-10";
		
		
		// 单表合计校验设置
		singleTotalMap.put("1-1", new String[]{"B12", "C12", "D11"});
	}

}
