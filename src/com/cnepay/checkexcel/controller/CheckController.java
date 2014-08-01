package com.cnepay.checkexcel.controller;

import java.io.File;
import java.util.List;

import com.cnepay.checkexcel.report.BankSingleExcelReport;
import com.cnepay.checkexcel.report.BaseExcelReport;
import com.cnepay.checkexcel.report.CheckResultMessage;
import com.cnepay.checkexcel.report.PaySingleExcelReport;
import com.cnepay.checkexcel.report.PayTotalExcelReport;
import com.cnepay.checkexcel.report.UnknownExcelReport;

public class CheckController {
	
	private BaseExcelReport report;
	
	public CheckController(File file) {
		String fileType = file.getName().split("\\.")[0].split("_")[0];
		int fileNameSectionNumber = file.getName().split("\\.")[0].split("_").length;
		
		if (fileType.equals(BaseExcelReport.TYPE_A) && fileNameSectionNumber == 4) {
			report = new PaySingleExcelReport(file);
		} else if (fileType.equals(BaseExcelReport.TYPE_A) && fileNameSectionNumber == 3) {
			report = new PayTotalExcelReport(file);
		} else if (fileType.equals(BaseExcelReport.TYPE_B) && fileNameSectionNumber == 4) {
			report = new BankSingleExcelReport(file);
		} else {
			report = new UnknownExcelReport(file);
		}
	}
	
	public CheckResultMessage checkFileName() {
		return report.checkFileName();
	}
	
	public CheckResultMessage checkSheetFormat() {
		return report.checkSheetFormat();
	}
	
	public CheckResultMessage checkDayofMonth() {
		return report.checkDayofMonth();
	}
	
	public CheckResultMessage checkSingleTotal() {
		return report.checkSingleTotal();
	}
	
	public List<CheckResultMessage> getMessageList() {
		return report.getMessageList();
	}
	
	public void close() {
		report.close();
	}
}
