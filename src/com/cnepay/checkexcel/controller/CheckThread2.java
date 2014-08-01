package com.cnepay.checkexcel.controller;

import java.io.File;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import com.cnepay.checkexcel.report.CheckResultMessage;
import com.cnepay.checkexcel.ui.Main1;

/*
 * 勾稽关系校验
 */
public class CheckThread2 implements Runnable {

	private int day;
	private String fileNameSections[];
	private String fileName;

	@Override
	public void run() {
		Main1.list.add(new CheckResultMessage("开始校验勾稽关系",
				CheckResultMessage.SYSTEM));
		for (File file : Main1.files1) {
			fileName = file.getName();
			fileNameSections = fileName.split("\\.")[0].split("_");
			String monthDate = fileNameSections[3];
			SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMM");
			try {
				Date date = dateFormat.parse(monthDate);
				Calendar cal = Calendar.getInstance();
				cal.setTime(date);
				day = cal.getActualMaximum(Calendar.DAY_OF_MONTH);

			} catch (ParseException e) {
				e.printStackTrace();
			}

			CheckController2 checkController;

			try {

				checkController = new CheckController2(file);
				for (int i = 1; i <= day; i++) {
					Main1.list.add(checkController.checkF(i));
					Main1.list.add(checkController.checkG(i));
				}

			} catch (Exception e) {

				e.printStackTrace();
			}
			try {
				checkController = new CheckController2(file);
				for (int i = 1; i <= day; i++) {
					Main1.list.add(checkController.checkF02(i));
					Main1.list.add(checkController.checkF04(i));
					Main1.list.add(checkController.checkF07(i));
					Main1.list.add(checkController.checkF05(i));
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
			try {
				checkController = new CheckController2(file);
				for (int i = 1; i <= day; i++) {
					Main1.list.add(checkController.checkJ01(i));
					Main1.list.add(checkController.checkJ02(i));
					Main1.list.add(checkController.checkJ03(i));
					Main1.list.add(checkController.checkJ04(i));
				}
			} catch (Exception e) {
			}
		}

		// Main1.list.add();
		// 检查汇总表勾稽
		try {
			CheckController2 checController2 = new CheckController2(
					Main1.files2);
			for (int i = 1; i <= day; i++) {
				Main1.list.add(checController2.check1(i));
				Main1.list.add(checController2.check2(i));
				Main1.list.add(checController2.check3(i));
				Main1.list.add(checController2.check4(i));
			}
		} catch (Exception e1) {
			e1.printStackTrace();
		}

		try {
			CheckController2 checController2 = new CheckController2(
					Main1.files2);
			for (int i = 2; i <= day; i++) {
				Main1.list.add(checController2.check5(i));
				Main1.list.add(checController2.check9(i));
				Main1.list.add(checController2.check11(i));
				Main1.list.add(checController2.check17(i));
				Main1.list.add(checController2.check18(i));
				Main1.list.add(checController2.check19(i));
				Main1.list.add(checController2.check20(i));
			}
			for (int i = 1; i <= day; i++) {
				Main1.list.add(checController2.check6(i));
				Main1.list.add(checController2.check7(i));
				Main1.list.add(checController2.checkE04(i));
				Main1.list.add(checController2.checkE05(i));
				Main1.list.add(checController2.check8(i));
			}
			for (int i = 1; i <= day; i++) {
				Main1.list.add(checController2.checkF01(i));
				Main1.list.add(checController2.checkF08(i));
				Main1.list.add(checController2.checkG09(i));
				Main1.list.add(checController2.checkG11(i));
				Main1.list.add(checController2.checkG12(i));
				Main1.list.add(checController2.checkTF02(i));
				Main1.list.add(checController2.checkTF07(i));
				Main1.list.add(checController2.checkTF04(i));
				Main1.list.add(checController2.checkTF05(i));
			}
			for (int i = 1; i <= day; i++) {
				Main1.list.add(checController2.checkF1(i));
				Main1.list.add(checController2.checkG1(i));
			}
			for (int i = 1; i <= day; i++) {
				Main1.list.add(checController2.check10(i));
				Main1.list.add(checController2.check12(i));
			}
			for (int i = 1; i <= day; i++) {
				Main1.list.add(checController2.check13(i));
				Main1.list.add(checController2.check14(i));
				Main1.list.add(checController2.check15(i));
				Main1.list.add(checController2.check16(i));
			}
			for (int i = 1; i <= day; i++) {
				Main1.list.add(checController2.checkL1(i));
				Main1.list.add(checController2.checkL2(i));
				Main1.list.add(checController2.checkL4(i));
				Main1.list.add(checController2.checkL5(i));
				Main1.list.add(checController2.checkL6(i));
				Main1.list.add(checController2.checkL7(i));
			}
			for (int i = 1; i <= day; i++) {
				Main1.list.add(checController2.check21(i));
				Main1.list.add(checController2.check22(i));
				Main1.list.add(checController2.check23(i));
				Main1.list.add(checController2.check24(i));
				Main1.list.add(checController2.check25(i));
				Main1.list.add(checController2.check26(i));
				Main1.list.add(checController2.check27(i));
				Main1.list.add(checController2.check28(i));
			}
			for (int i = 2; i <= day; i++) {
				Main1.list.add(checController2.check29(i));
				Main1.list.add(checController2.check30(i));
				Main1.list.add(checController2.check31(i));
				Main1.list.add(checController2.check32(i));
				Main1.list.add(checController2.check33(i));
				Main1.list.add(checController2.check34(i));
			}
		} catch (Exception e) {

		}
		Main1.list.add(new CheckResultMessage("勾稽关系校验完毕！",
				CheckResultMessage.SYSTEM));

		while (Main1.list.size() > 0) {
			// 需要sleep
			try {
				Thread.sleep(10);
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

		Main1.isChecking = false;
		Main1.freshUI();

	}

}
