package com.cnepay.checkexcel.report;

import java.io.File;

public class UnknownExcelReport extends BaseExcelReport {

	public UnknownExcelReport(File file) {
		super(file);
	}

	@Override
	public CheckResultMessage checkFileName() {
		return errorNameFormat();
	}

}
