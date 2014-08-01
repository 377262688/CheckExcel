package com.cnepay.checkexcel.controller;

import java.io.File;
import java.util.List;

import com.cnepay.checkexcel.report.BankSingleExcelReport1;
import com.cnepay.checkexcel.report.BaseReport;
import com.cnepay.checkexcel.report.CheckReport1;
import com.cnepay.checkexcel.report.CheckResultMessage;
import com.cnepay.checkexcel.report.PaySingleExcelReport1;
import com.cnepay.checkexcel.report.PayTotalExcelReport1;
import com.cnepay.checkexcel.report.UnknownExcelReport1;

public class CheckController1 {
	private BaseReport report;
	private CheckReport1 checkReport1;

	public CheckController1(File file) {
		String fileType = file.getName().split("\\.")[0].split("_")[0];
		int fileNameSectionNumber = file.getName().split("\\.")[0].split("_").length;
		try {
			checkReport1 = new CheckReport1(file);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		if (fileType.equals(BaseReport.TYPE_A) && fileNameSectionNumber == 4) {
			report = new PaySingleExcelReport1(file);

		} else if (fileType.equals(BaseReport.TYPE_A)
				&& fileNameSectionNumber == 3) {
			report = new PayTotalExcelReport1(file);
		} else if (fileType.equals(BaseReport.TYPE_B)
				&& fileNameSectionNumber == 4) {
			report = new BankSingleExcelReport1(file);
		} else {
			report = new UnknownExcelReport1(file);
		}
	}
	public int getMaxDay(){
		return report.maxDay;
	}
	public CheckResultMessage checkFileName() {
		return report.checkFileName();
	}

	public CheckResultMessage checkSheetFormat() {
		return report.checkSheetFormat();
	}

	public CheckResultMessage checkDayofMonth() {
		return checkReport1.checkDayofMonth();
	}

	public CheckResultMessage checkDayofMonth1() {
		return checkReport1.checkDayofMonth1();
	}

	public CheckResultMessage checkDayofMonth2() {
		return checkReport1.checkDayofMonth2();
	}
	public CheckResultMessage checkDayofMonth3() {
		return checkReport1.checkDayofMonth3();
	}
	public CheckResultMessage checkDayofMonth4() {
		return checkReport1.checkDayofMonth4();
	}
	public CheckResultMessage checkDayofMonth5() {
		return checkReport1.checkDayofMonth5();
	}
	public CheckResultMessage checkDayofMonth6() {
		return checkReport1.checkDayofMonth6();
	}

	public List<CheckResultMessage> getMessageList() {
		return report.getMessageList();
	}
	
	public CheckResultMessage checkTotal1(int i){
		return checkReport1.checkTotal1(i);
	}
	public CheckResultMessage checkTotal2(int i){
		return checkReport1.checkTotal2(i);
	}
	public CheckResultMessage checkTotal4(int i){
		return checkReport1.checkTotal4(i);
	}
	public CheckResultMessage checkSum1(int i,int j){
		return checkReport1.checkSum1(i,j);
	}
	public CheckResultMessage checkSum3(int i,int j){
		return checkReport1.checkSum3(i,j);
	}
	public CheckResultMessage checkSum4(int i,int j){
		return checkReport1.checkSum4(i,j);
	}
	public CheckResultMessage checkSum5(int i,int j){
		return checkReport1.checkSum5(i,j);
	}
}
