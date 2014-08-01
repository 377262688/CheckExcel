package com.cnepay.checkexcel.report;

import java.io.File;

/**
 * 支付机构汇总报表
 * - 支付机构汇总表的文件名为：A_支付机构名_期数.xls
 * - 支付机构汇总表只有1张，包括13个sheet，分别为1-1至1-13
 * - 逐表格逐项核对1-1中的合计数是否为本月所有日期的该项数据之和。逐项核对1-2、1-4中的合计数是否为本月所有日期的该项数据之和
 * 
 */
public class PayTotalExcelReport1 extends BaseReport {

	public PayTotalExcelReport1(File file) {
		super(file);

		// 命名以A开头
		this.fileType = BaseExcelReport.TYPE_A;
		// 命名使用下划线分为3节
		this.fileNameSectionsNumber = 3;
		// 报表类型
		this.reportName = "支付机构汇总报表";
		
		// Sheet名称
		this.sheetNames = new String[13];
		for(int i = 0; i < 13; i++) {
			sheetNames[i] = "1-" + (i+1); 
		}
		
		// 日期校验设置
		daysMap.put("1-1", new DayCell("A11", DayCell.COL_DIRECTION));
		daysMap.put("1-2", new DayCell("A9", DayCell.COL_DIRECTION));
		daysMap.put("1-3", new DayCell("E4", DayCell.ROW_DIRECTION));
		daysMap.put("1-4", new DayCell("A6", DayCell.COL_DIRECTION));
		daysMap.put("1-5", new DayCell("A7", DayCell.COL_DIRECTION));
		daysMap.put("1-6", new DayCell("C4", DayCell.ROW_DIRECTION));
		daysMap.put("1-7", new DayCell("A6", DayCell.COL_DIRECTION));
		daysMap.put("1-8", new DayCell("A6", DayCell.COL_DIRECTION));
		daysMap.put("1-9", new DayCell("C5", DayCell.ROW_DIRECTION));
		daysMap.put("1-10", new DayCell("A9", DayCell.COL_DIRECTION));
		daysMap.put("1-11", new DayCell("C4", DayCell.ROW_DIRECTION));
		daysMap.put("1-12", new DayCell("C4", DayCell.ROW_DIRECTION));
		daysMap.put("1-13", new DayCell("A7", DayCell.COL_DIRECTION));
		
		// 单表合计校验设置
		singleTotalMap.put("1-1", new String[]{"B12"});
		singleTotalMap.put("1-2", new String[]{"B10", "C10"});
		singleTotalMap.put("1-4", new String[]{"B7", "C7", "D7"});
	}

}
