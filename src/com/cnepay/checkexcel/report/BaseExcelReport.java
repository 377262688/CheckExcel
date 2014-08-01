package com.cnepay.checkexcel.report;

import java.io.File;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jxl.Cell;
import jxl.CellReferenceHelper;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.Workbook;




public abstract class BaseExcelReport {
	
	protected File file;
	protected List<CheckResultMessage> messageList = new ArrayList<CheckResultMessage>();

	protected String fileName;
	protected String fileNameSections[];
	protected String monthDate;
	protected Workbook workbook;
	
	
	protected String fileType;
	public static final String TYPE_A = "A";
	public static final String TYPE_B = "B";
	
	protected int fileNameSectionsNumber;
	
	protected String reportName;
	protected String sheetNames[];
	
	// 根据期数获取本月最后一天和上月最后一天
	protected int maxDay = 0;
	protected int lastMaxDay = 0;	
	
	// 日期检测设置
	protected Map<String, DayCell> daysMap = new HashMap<String, DayCell>();
	
	// 单表统计设置
	protected Map<String, String[]> singleTotalMap = new HashMap<String, String[]>();
	
	
	public BaseExcelReport(File file) {
		this.file = file;
		fileName = file.getName();
		fileNameSections = fileName.split("\\.")[0].split("_");
		
		try {
			workbook = Workbook.getWorkbook(file);
		} catch (Exception e) {
			e.printStackTrace();
			workbook = null;
		}
		
	}
	
	public CheckResultMessage pass(String message) {
		return new CheckResultMessage(message);
	}
	
	public CheckResultMessage error(String message) {
		return new CheckResultMessage(message, CheckResultMessage.CHECK_ERROR);
	}
	
	public CheckResultMessage errorNameFormat() {
		return new CheckResultMessage("该报表命名不规范：" + fileName, CheckResultMessage.CHECK_ERROR);
	}
	
	public CheckResultMessage errorGetWorkbook() {
		return new CheckResultMessage("读取该Excel文件失败：" + fileName, CheckResultMessage.CHECK_ERROR);
	}	

	/**
	 * 检查文件命名规范
	 * @return
	 */
	public CheckResultMessage checkFileName() {
		if (fileNameSections.length != fileNameSectionsNumber || ! fileNameSections[0].equals(fileType)) {
			return errorNameFormat();
		}
		
		if (fileNameSections.length == 4) {
			// 记录期数
			monthDate = fileNameSections[3];			
		} else if (fileNameSections.length == 3) {
			// 记录期数
			monthDate = fileNameSections[2];
		} else {
			return errorNameFormat();
		}
		
		return pass(reportName + " <" + fileName + "> 文件命名校验正确");
	}
	
	/**
	 * 检查Sheet的数量和命名
	 * @return
	 */
	public CheckResultMessage checkSheetFormat() {
		if (workbook == null) {
			return errorGetWorkbook();
		}
		
		for (String sheetName : sheetNames) {
			Sheet sheet = workbook.getSheet(sheetName);
			if (sheet == null) {
				return error(reportName + " 缺少Sheet " + sheetName);
			}
		}
		
		return pass(reportName + " <" + fileName + "> Sheet格式校验正确");
	}
	
	/**
	 * 根据期数核对每个表格中的日期是否正确（如期数为201302，则表格中的天数应为28，上月最后一天应为31）
	 * @return
	 */
	public CheckResultMessage checkDayofMonth() {
		if (workbook == null) {
			return errorGetWorkbook();
		}
		if (this.daysMap.size() < 1) {
			return null;
		}
		
		// clear message list
		messageList.clear();
		
		SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMM");
		
		try {
			Date date = dateFormat.parse(monthDate);
			Calendar cal = Calendar.getInstance();
			cal.setTime(date);
			
			maxDay = cal.getActualMaximum(Calendar.DAY_OF_MONTH);
			cal.add(Calendar.MONTH, -1);
			lastMaxDay = cal.getActualMaximum(Calendar.DAY_OF_MONTH);
			//System.out.println("maxDay=" + maxDay + ", lastMaxDay=" + lastMaxDay);
			
		} catch (ParseException e) {
			e.printStackTrace();
			return error(fileName + " 期数格式不正确！应为yyyyMM");
		}

		// 遍历检查
		for (String sheetName : sheetNames) {
			
			Sheet sheet = workbook.getSheet(sheetName);
			if (sheet == null) {
				continue;
			}
			
			DayCell days = daysMap.get(sheetName);
			if (days == null) {
				System.err.println("缺少days定义，sheetName=" + sheetName);
				continue;
			}
			
			if ( ! (lastMaxDay + "日").equals(sheet.getCell(days.getStartCol(), days.getStartRow()).getContents().trim())) {
				messageList.add(error("日期设置错误！<" + fileName + ">"
						+ " Sheet：" + sheetName
						+ ", 单元格" + CellReferenceHelper.getCellReference(days.getStartCol(), days.getStartRow())
						+ "设置为" + sheet.getCell(days.getStartCol(), days.getStartRow()).getContents()
						+ ", 应设为上月最后日期：" + lastMaxDay + "日"));
			}
			
			for (int i = 1; i <= maxDay; i++) {
				Cell cell = null;
				
				if (days.getDirection() == DayCell.COL_DIRECTION) {
					cell = sheet.getCell(days.getStartCol(), days.getStartRow() + i);
				} else if (days.getDirection() == DayCell.ROW_DIRECTION) {
					cell = sheet.getCell(days.getStartCol() + i, days.getStartRow());
				}
				
				if ( ! (i + "日").equals(cell.getContents().trim())) {
					messageList.add(error("日期设置错误！<" + fileName + ">" 
							+ " Sheet：" + sheetName
							+ ", 单元格" + CellReferenceHelper.getCellReference(cell.getColumn(), cell.getRow())
							+ "设置为" + cell.getContents()
							+ ", 应设为日期：" + i + "日"));
				}
			}
			
		}
		
		if (messageList.size() > 0) {
			return null;
		}
		
		return pass(reportName + " <" + fileName + "> 日期设置校验正确");
	}
	
	/**
	 * 检查单表每月统计
	 * 逐表格逐项核对1-1中的合计数是否为本月所有日期的该项数据之和。逐项核对1-2、1-4中的合计数是否为本月所有日期的该项数据之和
	 * @return
	 */
	public CheckResultMessage checkSingleTotal() {
		if (workbook == null) {
			return errorGetWorkbook();
		}
		if (this.singleTotalMap.size() < 1) {
			return null;
		}

		// clear message list
		messageList.clear();
		
		for (String sheetName : this.singleTotalMap.keySet()) {
				
			Sheet sheet = workbook.getSheet(sheetName); 
			if (sheet == null) {
				return null;
			}
			
			String countStartCells[] = singleTotalMap.get(sheetName);
			for (String countStartCell : countStartCells) {
				
				double totalCount = 0;
				
				int col = CellReferenceHelper.getColumn(countStartCell);
				int row = CellReferenceHelper.getRow(countStartCell);
				
				int rowIndex = row;
				for (rowIndex = row; rowIndex <= row + maxDay; rowIndex++) {
					
					Cell contentCell = sheet.getCell(col, rowIndex);
					
					if (contentCell == null || ! (contentCell instanceof NumberCell)) {
						messageList.add(error("单元格没有数据！<" + fileName + ">" 
								+ " Sheet：" + sheetName
								+ ", 单元格" + CellReferenceHelper.getCellReference(col, rowIndex)));
						continue;
					}
					
					totalCount = totalCount + ((NumberCell)sheet.getCell(col, rowIndex)).getValue();
				}
				
				Cell totalCell = sheet.getCell(col, rowIndex);
				if (totalCell == null || ! (totalCell instanceof NumberCell)) {
					messageList.add(error("合计单元格没有数据！<" + fileName + ">" 
							+ " Sheet：" + sheetName
							+ ", 单元格" + CellReferenceHelper.getCellReference(col, rowIndex)));
					continue;
				}
				
				double total = ((NumberCell)sheet.getCell(col, rowIndex)).getValue();
				if (totalCount != total) {
					return error("合计数据不正确！<" + fileName + ">" 
							+ " Sheet：" + sheetName
							+ ", 单元格" + CellReferenceHelper.getCellReference(col, rowIndex)
							+ ", 统计数据应为：" + totalCount
							+ ", 实际为：" + total);
				}
			}
		}
		
		if (messageList.size() > 0) {
			return null;
		}
		
		return pass(reportName + " <" + fileName + "> 单表合计数据校验正确");
	}
	
	public void close() {
		if (workbook != null) {
			workbook.close();
		}
	}
	
	public File getFile() {
		return file;
	}

	public void setFile(File file) {
		this.file = file;
	}
	
	public List<CheckResultMessage> getMessageList() {
		return messageList;
	}
	
}
