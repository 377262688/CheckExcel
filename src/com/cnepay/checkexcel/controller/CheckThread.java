package com.cnepay.checkexcel.controller;

import java.io.File;

import com.cnepay.checkexcel.report.CheckResultMessage;
import com.cnepay.checkexcel.report.StatData;
import com.cnepay.checkexcel.ui.MainFrame;

/**
 * 显示校验结果的线程
 *
 */
public class CheckThread implements Runnable {

	@Override
	public void run() {
		
		MainFrame.list.add(new CheckResultMessage("开始校验", CheckResultMessage.SYSTEM));
		
		StatData.init();
		
		for (File file : MainFrame.files) {
			CheckController checkController = new CheckController(file);
			// 校验文件命名
			MainFrame.list.add(checkController.checkFileName());
			// 校验表格格式
			MainFrame.list.add(checkController.checkSheetFormat());
			// 校验期数日期
			CheckResultMessage resultDayofMonth = checkController.checkDayofMonth();
			if (resultDayofMonth != null) {
				MainFrame.list.add(resultDayofMonth);
			}
			if (checkController.getMessageList() != null) {
				MainFrame.list.addAll(checkController.getMessageList());
			}
			// 校验单表合计数值
			CheckResultMessage resultSingleTotal = checkController.checkSingleTotal();
			if (resultSingleTotal != null) {
				MainFrame.list.add(resultSingleTotal);
			}
			if (checkController.getMessageList() != null) {
				MainFrame.list.addAll(checkController.getMessageList());
			}			
			
			checkController.close();
		}

		MainFrame.list.add(new CheckResultMessage("校验完毕！", CheckResultMessage.SYSTEM));
		
		while (MainFrame.list.size() > 0) {
			// 需要sleep
			try {
				Thread.sleep(10);
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		
		MainFrame.isChecking = false;
		MainFrame.freshUI();
	}

}
