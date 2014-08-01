package com.cnepay.checkexcel.ui;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FileFilter;
import java.util.Arrays;
import java.util.List;
import java.util.ArrayList;

import javax.swing.*;

import com.cnepay.checkexcel.controller.CheckThread;
import com.cnepay.checkexcel.report.CheckResultMessage;
import com.cnepay.checkexcel.report.StatData;

public class MainFrame {

	// 报表位置
	public static String path;
	// Excel文件集合
	public static List<File> files = new ArrayList<File>();
	// 结果消息队列
	public static List<CheckResultMessage> list = new ArrayList<CheckResultMessage>();
	// 校验状态
	public static boolean isChecking;
	
	
	// 主框架
	public static JFrame frame;
	// 消息面板
	public static MessagePanel panelMessage;
	// 打开按钮
	public static JButton btnOpen;
	// 校验按钮
	public static JButton btnCheck;
	// 状态显示
	public static JLabel lblStatus;
	
	
	public MainFrame() {
		
		try {
			UIManager.setLookAndFeel(UIManager.getCrossPlatformLookAndFeelClassName());
		} catch (Exception e) {
			e.printStackTrace();
		}		
		
		frame = new JFrame("备付金报表核对校验工具");
		Container contentPane = frame.getContentPane();
		contentPane.setBackground(Color.WHITE);

		// 操作面板位置
		JPanel panelButton = new JPanel();
		panelButton.setBackground(Color.LIGHT_GRAY);
		panelButton.setLayout(new FlowLayout());

		btnOpen = new JButton("选择报表文件夹");
		JLabel lblBlank = new JLabel("  ");
		btnCheck = new JButton("开始校验");

		panelButton.add(btnOpen);
		panelButton.add(lblBlank);
		panelButton.add(btnCheck);
		
		contentPane.add(panelButton, BorderLayout.NORTH);

		// 消息面板位置
		panelMessage = new MessagePanel("欢迎使用备付金报表核对工具！");
		contentPane.add(panelMessage, BorderLayout.CENTER);
		
		// 状态面板位置
		JPanel panelStatus = new JPanel();
		panelStatus.setBackground(Color.LIGHT_GRAY);
		panelStatus.setLayout(new FlowLayout());
		
		lblStatus = new JLabel();
		panelStatus.add(lblStatus);
		
		contentPane.add(panelStatus, BorderLayout.SOUTH);
//		contentPane.setComponentZOrder(panelStatus, contentPane.getComponentZOrder(panelStatus) - 2);
		
		// 点击【选择报表文件夹】按钮
		btnOpen.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent arg0) {
				// 清空files存储
				files.clear();
				
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				fileChooser.setDialogTitle("请选择备付金报表存放的文件夹");
				int result = fileChooser.showOpenDialog(frame);
				if (result == JFileChooser.APPROVE_OPTION) {
					
					File dir = fileChooser.getSelectedFile();
					path = dir.getAbsolutePath();
					panelMessage.addSystemMessage("打开报表存放位置：" + path + "，可以开始校验！");
					
					freshStatus("已就绪");
					
					File[] files = dir.listFiles(new FileFilter() {
						@Override
						public boolean accept(File file) {
							if (file.isFile() && 
									(file.getName().endsWith(".xls") || file.getName().endsWith(".xlsx"))) {
								return true;
							}
							return false;
						}
					});
					
					if (files.length <= 0) {
						panelMessage.addWarnMessage("没有找到Excel报表文件！请检查文件夹路径");
						return;
					}
					
					// 存储File对象数组
					MainFrame.files.addAll(Arrays.asList(files));
				}
			}
		});
		
		// 点击【开始校验】按钮
		btnCheck.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent arg0) {
				if (MainFrame.files == null || MainFrame.files.size() <= 0) {
					panelMessage.addWarnMessage("没有找到Excel报表文件！请检查文件夹路径");
					return;
				}
				
				isChecking = true;
				freshUI();
				
				Thread threadCheck = new Thread(new CheckThread());
				threadCheck.start();
			}
		});
		

		// 设置窗口事件
		frame.addWindowListener(new WindowAdapter() {

			// 退出
			@Override
			public void windowClosing(WindowEvent e) {
				// 正在检测过程中，提示警告
				if (isChecking) {
					int flag = JOptionPane.showConfirmDialog(frame.getContentPane(),
							"正在校验过程中，真的要退出吗？",
							"警告",
							JOptionPane.YES_NO_OPTION,
							JOptionPane.WARNING_MESSAGE);
					
					if (flag == JOptionPane.OK_OPTION) {
						System.exit(0);
					}
					
				} else {
					System.exit(0);
				}
			}		
		});

	}
	
	public void show() {
		// 根据内容调整窗口大小
		//frame.pack();
		//Dimension d = frame.getSize();
		//frame.setSize((int)d.getWidth() + 10, (int)d.getHeight() + 15);
		frame.setSize(800, 600);
		// 居中
		frame.setLocationRelativeTo(null);
		// 显示窗口
		frame.setVisible(true);
		// 取消默认关闭窗口
		frame.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
		
		// 启动消息显示
		Thread messageThread = new Thread(panelMessage);
		messageThread.start();
		
		freshStatus("已就绪");
	}
	
	// 刷新界面状态
	public static void freshUI() {
		if (isChecking) {
			btnOpen.setEnabled(false);
			btnCheck.setEnabled(false);
			
			freshStatus("检测中");
			
		} else {
			btnOpen.setEnabled(true);
			btnCheck.setEnabled(true);
			
			freshStatus("已就绪");
		}
	}
	
	// 刷新状态栏
	public static void freshStatus(String status) {
		
		String pathStatus = "";
		if (path == null || path.isEmpty()) {
			pathStatus = "【没有选中的报表文件夹】";
		} else {
			pathStatus = "【当前报表位置：" + path + "】";
		}
		
		String statNumber = "";
		if (StatData.errorNumber > 0) {
			statNumber = "本次检测发现错误" + StatData.errorNumber + "个";
		}
		
		lblStatus.setText(pathStatus + "    " + status + "    " + statNumber);
	}
	
	public static void main(String args[]) {
		MainFrame mainFrame = new MainFrame();
		mainFrame.show();
	}

}