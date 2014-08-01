package com.cnepay.checkexcel.report;

import java.io.File;

public class UnknownExcelReport1 extends BaseReport {

	public UnknownExcelReport1(File file) {
		super(file);
	}
	@Override
	public CheckResultMessage checkFileName() {
		return errorNameFormat();
	}

}
