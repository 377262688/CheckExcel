package com.cnepay.checkexcel.controller;

import java.io.File;
import java.util.List;

import com.cnepay.checkexcel.report.CheckReport1;
import com.cnepay.checkexcel.report.CheckResultMessage;
import com.cnepay.checkexcel.report.StatData;
import com.cnepay.checkexcel.ui.Main1;

public class CheckThread1 implements Runnable {

	public List<CheckResultMessage> message;

	@Override
	public void run() {

		Main1.list
				.add(new CheckResultMessage("开始校验", CheckResultMessage.SYSTEM));
		StatData.init();
		for (File file : Main1.files) {
			CheckController1 checkController = new CheckController1(file);
			// 校验文件命名
			Main1.list.add(checkController.checkFileName());
			// 校验表格格式
			Main1.list.add(checkController.checkSheetFormat());
		}
		for (File file : Main1.files1) {
			CheckController1 checkController = new CheckController1(file);
			// 校验期数日期
			Main1.list.add(checkController.checkDayofMonth());
			Main1.list.add(checkController.checkDayofMonth1());
			Main1.list.add(checkController.checkDayofMonth2());
			Main1.list.add(checkController.checkDayofMonth3());
			Main1.list.add(checkController.checkDayofMonth4());
		}

		CheckController1 checkController1 = new CheckController1(Main1.files2);
		Main1.list.add(checkController1.checkDayofMonth5());
		Main1.list.add(checkController1.checkDayofMonth6());
		// 检查单支付表1-1合计
		for (File file : Main1.files1) {
			CheckController1 checkController = new CheckController1(file);
			for (int i = 0; i < 14; i++) {
				Main1.list.add(checkController.checkTotal1(i));
			}
		}
		// 检查汇总表合计1-2,1-4
		CheckController1 checkController = new CheckController1(Main1.files2);
		for (int i = 0; i < 9; i++) {
			Main1.list.add(checkController.checkTotal2(i));
		}
		for (int i = 0; i < 3; i++) {
			Main1.list.add(checkController.checkTotal4(i));
		}
		// 检查所有单表1-1之和和汇总表格对应数据
		CheckController1 checkController2 = new CheckController1(Main1.files2);
		for (int i = 0; i < 9; i++) {
			for (int j = 1; j <= 32; j++) {
				Main1.list.add(checkController2.checkSum1(i, j));
			}
		}
		for (int i = 0; i < 40; i++)
			for (int j = 1; j <= 31; j++) {
				Main1.list.add(checkController2.checkSum3(i, j));
			}
		for(int i=0;i<4;i++)
			for(int j=1;j<=32;j++){
				Main1.list.add(checkController2.checkSum4(i, j));
			}
		for(int i=0;i<24;i++){
			for(int j=1;j<=31;j++){
				Main1.list.add(checkController2.checkSum5(i, j));
			}
		}
		Main1.list.add(new CheckResultMessage("校验完毕！",
				CheckResultMessage.SYSTEM));

		while (Main1.list.size() > 0) {
			// 需要sleep
			try {
				Thread.sleep(10);
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
		}

		Main1.isChecking = false;
		Main1.freshUI();
	}

}
