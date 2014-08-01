package com.cnepay.checkexcel.report;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.cnepay.checkexcel.ui.Main1;

public abstract class BaseReport {

	protected File file;
	protected List<CheckResultMessage> messageList = new ArrayList<CheckResultMessage>();

	protected String fileName;
	protected String fileNameSections[];
	protected String monthDate;
	protected XSSFWorkbook workbook;
	protected XSSFWorkbook xWorkbook;
	protected XSSFSheet xSheet;
	protected XSSFRow xRow;
	protected XSSFCell xCell;
	protected HSSFWorkbook hWorkbook;
	protected HSSFSheet hSheet;
	protected HSSFRow hRow;
	protected HSSFCell hCell;

	protected String fileType;
	public static final String TYPE_A = "A";
	public static final String TYPE_B = "B";

	protected int fileNameSectionsNumber;

	protected String reportName;
	protected String sheetNames[];

	// 根据期数获取本月最后一天和上月最后一天
	public int maxDay = 0;
	protected int lastMaxDay = 0;


	// 单表统计设置
	protected Map<String, String[]> singleTotalMap = new HashMap<String, String[]>();

	public BaseReport(File file) {
		this.file = file;
		fileName = file.getName();
		fileNameSections = fileName.split("\\.")[0].split("_");

		try {
			if (checkVersion(file).equals("2003")) {
				InputStream in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
			} else {
				InputStream in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public CheckResultMessage pass(String message) {
		return new CheckResultMessage(message);
	}

	public CheckResultMessage error(String message) {
		return new CheckResultMessage(message, CheckResultMessage.CHECK_ERROR);
	}

	public CheckResultMessage errorNameFormat() {
		return new CheckResultMessage("该报表命名不规范：" + fileName,
				CheckResultMessage.CHECK_ERROR);
	}

	public CheckResultMessage errorGetWorkbook() {
		return new CheckResultMessage("读取该Excel文件失败：" + fileName,
				CheckResultMessage.CHECK_ERROR);
	}

	/**
	 * 检查文件命名规范
	 * 
	 * @return
	 */
	public CheckResultMessage checkFileName() {
		if (fileNameSections.length != fileNameSectionsNumber
				|| !fileNameSections[0].equals(fileType)) {
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
	 * 
	 * @return
	 */
	public CheckResultMessage checkSheetFormat() {
		if(hWorkbook == null && xWorkbook == null){
			return error("读取报表"+fileName+"失败");
		}
		if (checkVersion(file).equals("2003")) {
			for (String sheetName : sheetNames) {
				hSheet = hWorkbook.getSheet(sheetName);
				if (hSheet == null) {
					return error(reportName + " 缺少Sheet " + sheetName);
				}
			}
		} else {
			for (String sheetName : sheetNames) {
				xSheet = xWorkbook.getSheet(sheetName);
				if (xSheet == null) {
					return error(reportName + " 缺少Sheet " + sheetName);
				}
			}
		}

		return pass(reportName + " <" + fileName + "> Sheet格式校验正确");
	}

	/**
	 * 根据期数核对每个表格中的日期是否正确（如期数为201302，则表格中的天数应为28，上月最后一天应为31）
	 * 
	 * @return
	 */
	public CheckResultMessage checkDayofMonth() {

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

		} catch (ParseException e) {
			e.printStackTrace();
			return error(fileName + " 期数格式不正确！应为yyyyMM");
		}
		if (checkVersion(file).equals("2003")) {
			if (fileNameSections.length == 4) {
				int r1 = get(2, 2);
				int c1 = get(3, 2);
				if (!(getValueString(r1 + 1, c1 - 1, 0).equals(lastMaxDay + "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-1"
							+ ", 应设为上月最后日期：" + lastMaxDay + "日");
				}
			}
		} else {
			if (fileNameSections.length == 4) {
				int r1 = get(2, 2);
				int c1 = get(3, 2);
				if (!(getValueString1(r1 + 1, c1 - 1, 0).equals(lastMaxDay + "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-1"
							+ ", 应设为上月最后日期：" + lastMaxDay + "日");
				}
			}
		}
		// 遍历检查
		// for (String sheetName : sheetNames) {
		//
		// XSSFSheet sheet = workbook.getSheet(sheetName);
		// if (sheet == null) {
		// continue;
		// }
		//
		// DayCell days = daysMap.get(sheetName);
		// if (days == null) {
		// System.err.println("缺少days定义，sheetName=" + sheetName);
		// continue;
		// }
		// XSSFRow row = sheet.getRow(days.getStartCol());
		// XSSFCell cell = row.getCell(days.getStartRow());
		// if ( ! (lastMaxDay + "日").equals(cell.getStringCellValue().trim())) {
		// messageList.add(error("日期设置错误！<" + fileName + ">"
		// + " Sheet：" + sheetName
		// // + ", 单元格" +
		// CellReferenceHelper.getCellReference(days.getStartCol(),
		// days.getStartRow())
		// + "设置为" + cell.getStringCellValue()
		// + ", 应设为上月最后日期：" + lastMaxDay + "日"));
		// }
		//
		// for (int i = 1; i <= maxDay; i++) {
		// XSSFComment cell1 = null;
		//
		// if (days.getDirection() == DayCell.COL_DIRECTION) {
		// cell1 = sheet.getCellComment(days.getStartCol(), days.getStartRow() +
		// i);
		// } else if (days.getDirection() == DayCell.ROW_DIRECTION) {
		// cell1 = sheet.getCellComment(days.getStartCol() + i,
		// days.getStartRow());
		// }
		//
		// if ( ! (i + "日").equals(( cell1.getString()))) {
		// messageList.add(error("日期设置错误！<" + fileName + ">"
		// + " Sheet：" + sheetName
		// + ", 单元格" + CellReferenceHelper.getCellReference(cell.getColumn(),
		// cell.getRow())
		// + "设置为" + cell1.getContents()
		// + ", 应设为日期：" + i + "日"));
		// }
		// }
		//
		// }

		if (messageList.size() > 0) {
			return null;
		}

		return pass(reportName + " <" + fileName + "> 日期设置校验正确");
	}

	/**
	 * 检查单表每月统计 逐表格逐项核对1-1中的合计数是否为本月所有日期的该项数据之和。逐项核对1-2、1-4中的合计数是否为本月所有日期的该项数据之和
	 * 
	 * @return
	 */
	/*
	 * public CheckResultMessage checkSingleTotal() { if (workbook == null) {
	 * return errorGetWorkbook(); } if (this.singleTotalMap.size() < 1) { return
	 * null; }
	 * 
	 * // clear message list messageList.clear();
	 * 
	 * for (String sheetName : this.singleTotalMap.keySet()) { InputStream in =
	 * new FileInputStream(file); XSSFSheet sheet =
	 * workbook.getSheet(sheetName); if (sheet == null) { return null; }
	 * 
	 * String countStartCells[] = singleTotalMap.get(sheetName); for (String
	 * countStartCell : countStartCells) {
	 * 
	 * double totalCount = 0;
	 * 
	 * int col = CellReferenceHelper.getColumn(countStartCell); int row =
	 * CellReferenceHelper.getRow(countStartCell);
	 * 
	 * int rowIndex = row; for (rowIndex = row; rowIndex <= row + maxDay;
	 * rowIndex++) {
	 * 
	 * Cell contentCell = sheet.getCell(col, rowIndex);
	 * 
	 * if (contentCell == null || ! (contentCell instanceof NumberCell)) {
	 * messageList.add(error("单元格没有数据！<" + fileName + ">" + " Sheet：" +
	 * sheetName + ", 单元格" + CellReferenceHelper.getCellReference(col,
	 * rowIndex))); continue; }
	 * 
	 * totalCount = totalCount + ((NumberCell)sheet.getCell(col,
	 * rowIndex)).getValue(); }
	 * 
	 * Cell totalCell = sheet.getCell(col, rowIndex); if (totalCell == null || !
	 * (totalCell instanceof NumberCell)) { messageList.add(error("合计单元格没有数据！<"
	 * + fileName + ">" + " Sheet：" + sheetName + ", 单元格" +
	 * CellReferenceHelper.getCellReference(col, rowIndex))); continue; }
	 * 
	 * double total = ((NumberCell)sheet.getCell(col, rowIndex)).getValue(); if
	 * (totalCount != total) { return error("合计数据不正确！<" + fileName + ">" +
	 * " Sheet：" + sheetName + ", 单元格" +
	 * CellReferenceHelper.getCellReference(col, rowIndex) + ", 统计数据应为：" +
	 * totalCount + ", 实际为：" + total); } } }
	 * 
	 * if (messageList.size() > 0) { return null; }
	 * 
	 * return pass(reportName + " <" + fileName + "> 单表合计数据校验正确"); }
	 */

	// public void close() {
	// if (workbook != null) {
	// workbook.close();
	// }
	// }

	public File getFile() {
		return file;
	}

	public void setFile(File file) {
		this.file = file;
	}

	public List<CheckResultMessage> getMessageList() {
		return messageList;
	}

	/**
	 * 获取单个个格子的数据
	 * 
	 * @return
	 */
	public BigDecimal getValue(int row, int clums, int sheetat) {
		BigDecimal value = new BigDecimal(0);
		String s = null;
		try {
			hSheet = hWorkbook.getSheetAt(sheetat - 1);
			hRow = hSheet.getRow(row - 1);
			hCell = hRow.getCell(clums - 1);
			hCell.setCellType(Cell.CELL_TYPE_STRING);
			s = hCell.getStringCellValue();
			value = new BigDecimal(s).setScale(2, BigDecimal.ROUND_HALF_UP);
		} catch (Exception e) {
			value = new BigDecimal(0);
		}
		return value;
	}

	public String getValueString(int row, int clums, int sheetat) {
		String s = null;
		try {
			hSheet = hWorkbook.getSheetAt(sheetat - 1);
			hRow = hSheet.getRow(row - 1);
			hCell = hRow.getCell(clums - 1);
			hCell.setCellType(Cell.CELL_TYPE_STRING);
			s = hCell.getStringCellValue();
		} catch (Exception e) {

		}
		return s;
	}

	/**
	 * 获取单个个格子的数据
	 * 
	 * @return
	 */
	public BigDecimal getValue1(int row, int clums, int sheetat) {
		BigDecimal value = new BigDecimal(0);
		String s = null;
		try {
			XSSFSheet xSheet = xWorkbook.getSheetAt(sheetat);
			XSSFRow xRow = xSheet.getRow(row - 1);
			XSSFCell xCell = xRow.getCell(clums - 1);
			xCell.setCellType(Cell.CELL_TYPE_STRING);
			s = xCell.getStringCellValue();
			value = new BigDecimal(s).setScale(2, BigDecimal.ROUND_HALF_UP);

		} catch (Exception e) {
			value = new BigDecimal(0);
		}
		return value;
	}

	public String getValueString1(int row, int clums, int sheetat) {
		String s = null;
		try {
			XSSFSheet xSheet = xWorkbook.getSheetAt(sheetat);
			XSSFRow xRow = xSheet.getRow(row - 1);
			XSSFCell xCell = xRow.getCell(clums - 1);
			xCell.setCellType(Cell.CELL_TYPE_STRING);
			s = xCell.getStringCellValue();
		} catch (Exception e) {
		}
		return s;
	}

	/**
	 * 检查版本
	 * 
	 * @return 版本号
	 */
	public String checkVersion(File file) {
		String v = null;
		if (fileName.endsWith(".xls")) {
			v = "2003";
		} else if (fileName.endsWith(".xlsx")) {
			v = "2007";
		}
		return v;
	}

	/*
	 * 从配置文件中取得行号和列
	 */
	public int get(int row, int clums) {
		int i = 0;
		File f = new File("C:/config/config.xlsx");
		try {
			InputStream input = new FileInputStream(f);
			XSSFWorkbook xssfWorkbook = new XSSFWorkbook(input);
			XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
			XSSFRow xssfRow = xssfSheet.getRow(row - 1);
			XSSFCell xssfCell = xssfRow.getCell(clums - 1);
			xssfCell.setCellType(Cell.CELL_TYPE_STRING);
			String s = xssfCell.getStringCellValue();
			i = Integer.parseInt(s);
		} catch (Exception e) {

		}
		return i;
	}
}
