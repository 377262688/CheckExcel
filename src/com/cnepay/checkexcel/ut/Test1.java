package com.cnepay.checkexcel.ut;

//import static org.junit.Assert.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import jxl.biff.CellReferenceHelper;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class Test1 {
	private XSSFWorkbook xWorkbook;
	private File f;
	@Test
	public void sy(){
		f = new File("A_建设银行_4839204320_201403.xlsx");
		try {
			InputStream in = new FileInputStream(f);
			xWorkbook = new XSSFWorkbook(in);
		} catch (Exception e) {
		}
		System.out.println(getValueString1(13, 1, 0));
		System.out.println(getValueString1(44, 1, 0));
	}
	
	public BigDecimal getValue1(int row, int clums, int sheetat) {
		BigDecimal value = new BigDecimal(0);
		String s = null;
		try {
			XSSFSheet xSheet = xWorkbook.getSheetAt(sheetat);	
			XSSFRow xRow = xSheet.getRow(row - 1);	
			XSSFCell xCell = xRow.getCell(clums - 1);	
			xCell.setCellType(Cell.CELL_TYPE_STRING);	
			s = xCell.getStringCellValue();
			value=new BigDecimal(s).setScale(2, BigDecimal.ROUND_HALF_UP);
			
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




}
