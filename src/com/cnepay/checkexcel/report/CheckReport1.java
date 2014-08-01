package com.cnepay.checkexcel.report;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

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

public class CheckReport1 {
	protected File file;
	protected String fileName;
	protected String reportName;
	protected InputStream in;
	protected String version;
	protected XSSFWorkbook xWorkbook;
	protected XSSFSheet xSheet;
	protected XSSFRow xRow;
	protected XSSFCell xCell;
	protected HSSFWorkbook hWorkbook;
	protected HSSFSheet hSheet;
	protected HSSFRow hRow;
	protected HSSFCell hCell;
	protected String fileNameSections[];
	protected String monthDate;
	protected int maxDay;
	protected int lastMaxDay;
	protected List<CheckResultMessage> message;

	public CheckReport1(File file) throws Exception {
		this.file = file;
		fileName = file.getName();
		fileNameSections = fileName.split("\\.")[0].split("_");
		if (fileNameSections.length == 4) {
			// 记录期数
			monthDate = fileNameSections[3];
		} else if (fileNameSections.length == 3) {
			// 记录期数
			monthDate = fileNameSections[2];
		}
	}

	/*
	 * 检查单表1-1日期
	 */
	public CheckResultMessage checkDayofMonth() {
		int r1 = get(2, 2);
		int c1 = get(3, 2);
		date();
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (!(getValueString(r1 + 1, c1 - 1, 0)
						.equals(lastMaxDay + "日") && getValueString(
						r1 + 1 + maxDay, c1 - 1, 0).equals(maxDay + "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-1"
							+ ", 应设为上月最后日期：" + lastMaxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (!(getValueString1(r1 + 1, c1 - 1, 0).equals(
						lastMaxDay + "日") && (getValueString1(r1 + 1 + maxDay,
						c1 - 1, 0).equals(maxDay + "日")))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-1"
							+ ", 应设为上月最后日期：" + lastMaxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付支付机构单个账户表格：<" + fileName + ">Sheet:1-1日期核对正确");
	}

	/*
	 * 检查单表1-3期数
	 */
	public CheckResultMessage checkDayofMonth1() {
		int r1 = get(6, 2);
		int c1 = get(7, 2);
		date();
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (!(getValueString(r1 - 1, c1 + maxDay, 1).equals(maxDay
						+ "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-3"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (!(getValueString1(r1 - 1, c1 + maxDay, 1).equals(maxDay
						+ "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-3"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付支付机构单个账户表格：<" + fileName + ">Sheet:1-3日期核对正确");
	}

	/*
	 * 检查单表1-6期数
	 */
	public CheckResultMessage checkDayofMonth2() {
		int r1 = get(10, 2);
		int c1 = get(11, 2);
		date();
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (!(getValueString(r1 - 2, c1 + maxDay, 2).equals(maxDay
						+ "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-6"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (!(getValueString1(r1 - 2, c1 + maxDay, 2).equals(maxDay
						+ "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-6"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付支付机构单个账户表格：<" + fileName + ">Sheet:1-6日期核对正确");
	}

	/*
	 * 检查单表1-9期数
	 */
	public CheckResultMessage checkDayofMonth3() {
		int r1 = get(17, 2);
		int c1 = get(18, 2);
		date();
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (!(getValueString1(r1 - 1, c1 + 1 + maxDay, 3).equals(
						maxDay + "日") && (getValueString1(r1 - 1, c1 + 1, 3)
						.equals(lastMaxDay + "日")))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-9"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				System.out.println(getValueString1(r1 - 1, c1 + 1 + maxDay, 3));
				if (!(getValueString1(r1 - 1, c1 + 1 + maxDay, 3).equals(
						maxDay + "日") && (getValueString1(r1 - 1, c1 + 1, 3)
						.equals(lastMaxDay + "日")))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-9"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付支付机构单个账户表格：<" + fileName + ">Sheet:1-9日期核对正确");
	}

	/*
	 * 检查单表1-10期数
	 */
	public CheckResultMessage checkDayofMonth4() {
		int r1 = get(20, 2);
		int c1 = get(21, 2);
		date();
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (!(getValueString(r1 + maxDay, c1 - 1, 4).equals(maxDay
						+ "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-10"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (!(getValueString(r1 + maxDay, c1 - 1, 4).equals(maxDay
						+ "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-10"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付支付机构单个账户表格：<" + fileName + ">Sheet:1-10日期核对正确");
	}

	/*
	 * 检查汇总表格1-1日期
	 */
	public CheckResultMessage checkDayofMonth5() {
		int r1 = get(2, 5);
		int c1 = get(3, 5);
		date();
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (!(getValueString(r1 + maxDay, c1 - 1, 4).equals(maxDay
						+ "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-1"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (!(getValueString(r1 + maxDay, c1 - 1, 4).equals(maxDay
						+ "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-1"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付支付机构汇总表格：<" + fileName + ">Sheet:1-1日期核对正确");
	}

	/*
	 * 检查汇总表格1-2日期
	 */
	public CheckResultMessage checkDayofMonth6() {
		int r1 = get(6, 5);
		int c1 = get(7, 5);
		date();
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (!(getValueString(r1 + maxDay, c1 - 1, 4).equals(maxDay
						+ "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-2"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (!(getValueString(r1 + maxDay, c1 - 1, 4).equals(maxDay
						+ "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-2"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付支付机构汇总表格：<" + fileName + ">Sheet:1-2日期核对正确");
	}

	/*
	 * 检查汇总表格1-3日期
	 */
	public CheckResultMessage checkDayofMonth7() {
		int r1 = get(10, 5);
		int c1 = get(11, 5);
		date();
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (!(getValueString(r1 - 1, c1 + maxDay, 2).equals(maxDay
						+ "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-3"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (!(getValueString1(r1 - 1, c1 + maxDay, 2).equals(maxDay
						+ "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-3"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付支付机构汇总表格：<" + fileName + ">Sheet:1-3日期核对正确");
	}

	/*
	 * 检查汇总表格1-4日期
	 */
	public CheckResultMessage checkDayofMonth8() {
		int r1 = get(13, 5);
		int c1 = get(14, 5);
		date();
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (!(getValueString(r1 + maxDay, c1 - 1, 3).equals(maxDay
						+ "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-4"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (!(getValueString(r1 + maxDay, c1 - 1, 3).equals(maxDay
						+ "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-2"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付支付机构汇总表格：<" + fileName + ">Sheet:1-4日期核对正确");
	}

	/*
	 * 检查汇总表格1-5日期
	 */
	public CheckResultMessage checkDayofMonth9() {
		int r1 = get(13, 5);
		int c1 = get(14, 5);
		date();
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (!(getValueString(r1 + maxDay, c1 - 1, 3).equals(maxDay
						+ "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-4"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (!(getValueString(r1 + maxDay, c1 - 1, 3).equals(maxDay
						+ "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-2"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付支付机构汇总表格：<" + fileName + ">Sheet:1-5日期核对正确");
	}

	/*
	 * 检查汇总表格1-6日期
	 */
	public CheckResultMessage checkDayofMonth10() {
		int r1 = get(13, 5);
		int c1 = get(14, 5);
		date();
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (!(getValueString(r1 + maxDay, c1 - 1, 3).equals(maxDay
						+ "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-4"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (!(getValueString(r1 + maxDay, c1 - 1, 3).equals(maxDay
						+ "日"))) {
					return error("日期设置错误！<" + fileName + ">" + " Sheet：1-2"
							+ ", 应设为本月最后日期：" + maxDay + "日");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付支付机构汇总表格：<" + fileName + ">Sheet:1-6日期核对正确");
	}

	/*
	 * 检查1-1所有日期之和
	 */
	public CheckResultMessage checkTotal1(int i) {
		int r1 = get(2, 2);
		int c1 = get(3, 2);
		int r2 = get(4, 2);
		BigDecimal sum = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				for (int j = 1; j < r2 - r1; j++) {
					sum = sum.add(getValue(r1 + j, c1 + i, 0));
				}
				if (0 != getValue(r2, c1 + i, 0).compareTo(sum)) {
					return error("支付机构单个账户报表<" + fileName + ">sheet1-1:"
							+ "本月所有日期合计:A0" + (1 + i) + "错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				for (int j = 1; j < r2 - r1; j++) {
					sum = sum.add(getValue1(r1 + j, c1 + i, 0));
				}
				if (0 != getValue1(r2, c1 + i, 0).compareTo(sum)) {
					return error("支付机构单个账户报表<" + fileName + ">sheet1-1:"
							+ "本月所有日期合计:A0" + (1 + i) + "错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构单个账户报表<" + fileName + ">sheet1-1:" + "本月所有日期合计:A0"
				+ (1 + i) + "正确");
	}

	/*
	 * 检查1-2所有日期之和
	 */
	public CheckResultMessage checkTotal2(int i) {
		int r1 = get(6, 5);
		int c1 = get(7, 5);
		int r2 = get(8, 5);
		BigDecimal sum = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				for (int j = 1; j < r2 - r1; j++) {
					sum = sum.add(getValue(r1 + j, c1 + i, 1));
				}
				if (0 != getValue(r2, c1 + i, 1).compareTo(sum)) {
					return error("支付机构汇总报表<" + fileName + ">sheet1-2:"
							+ "本月所有日期合计:B0" + (1 + i) + "错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				for (int j = 1; j < r2 - r1; j++) {
					sum = sum.add(getValue1(r1 + j, c1 + i, 1));
				}
				if (0 != getValue1(r2, c1 + i, 1).compareTo(sum)) {
					return error("支付机构汇总报表<" + fileName + ">sheet1-2:"
							+ "本月所有日期合计:B0" + (1 + i) + "错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">sheet1-2:" + "本月所有日期合计:B0"
				+ (1 + i) + "正确");
	}

	/*
	 * 检查1-4所有日期之和
	 */
	public CheckResultMessage checkTotal4(int i) {
		int r1 = get(13, 5);
		int c1 = get(14, 5);
		int r2 = get(15, 5);
		BigDecimal sum = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				for (int j = 1; j < r2 - r1; j++) {
					sum = sum.add(getValue(r1 + j, c1 + i, 3));
				}
				if (0 != getValue(r2, c1 + i, 3).compareTo(sum)) {
					return error("支付机构汇总报表<" + fileName + ">sheet1-4:"
							+ "本月所有日期合计:D0" + (1 + i) + "错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				for (int j = 1; j < r2 - r1; j++) {
					sum = sum.add(getValue1(r1 + j, c1 + i, 3));
				}
				if (0 != getValue1(r2, c1 + i, 3).compareTo(sum)) {
					return error("支付机构汇总报表<" + fileName + ">sheet1-4:"
							+ "本月所有日期合计:D0" + (1 + i) + "错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">sheet1-4:" + "本月所有日期合计:D0"
				+ (1 + i) + "正确");
	}
	/*
	 * 检查1-1每个格子的相加的和
	 */
	public CheckResultMessage checkSum1(int i, int j) {
		int r1 = get(2, 2);
		int c1 = get(3, 2);
		int tr1 = get(2, 5);
		int tc1 = get(3, 5);
		BigDecimal sum = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			for (File f : Main1.files1) {
				try {
					in = new FileInputStream(f);
					hWorkbook = new HSSFWorkbook(in);
					sum = sum.add(getValue(r1 + j, c1 + i, 0));
					in.close();
				} catch (Exception e) {
				}
			}
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(tr1 + j, tc1 + i, 0).compareTo(sum)) {
					return error("支付机构单个账户表格sheet1-1的数据之和A0" + (i + 1) + "列"
							+ (r1 + j) + "行" + "出错");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			for (File f : Main1.files1) {
				try {
					in = new FileInputStream(f);
					xWorkbook = new XSSFWorkbook(in);
					sum = sum.add(getValue1(r1 + j, c1 + i, 0));
					in.close();
				} catch (Exception e) {
				}
			}
			
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue1(tr1 + j, tc1 + i, 0).compareTo(sum)) {
					return error("支付机构单个账户表格sheet1-1的数据之和A0" + (i + 1) + "列"
							+ (r1 + j) + "行" + "出错");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构单个账户表格sheet1-1的数据之和A0" + (i + 1) + "列" + (r1 + j)
				+ "行" + "正确");
	}
	/*
	 * 检查1-3每个格子的相加的和
	 */
	public CheckResultMessage checkSum2(int i, int j) {
		int r1 = get(5, 2);
		int c1 = get(6, 2);
		int tr1 = get(10, 5);
		int tc1 = get(11, 5);
		BigDecimal sum = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			for (File f : Main1.files1) {
				try {
					in = new FileInputStream(f);
					hWorkbook = new HSSFWorkbook(in);
					sum = sum.add(getValue(r1 + j, c1 + i, 2));
					in.close();
				} catch (Exception e) {
				}
			}
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(tr1 + j, tc1 + i, 0).compareTo(sum)) {
					return error("支付机构单个账户表格sheet1-1的数据之和A0" + (i + 1) + "列"
							+ (r1 + j) + "行" + "出错");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			for (File f : Main1.files1) {
				try {
					in = new FileInputStream(f);
					xWorkbook = new XSSFWorkbook(in);
					sum = sum.add(getValue1(r1 + j, c1 + i, 0));
					in.close();
				} catch (Exception e) {
				}
			}
			
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue1(tr1 + j, tc1 + i, 0).compareTo(sum)) {
					return error("支付机构单个账户表格sheet1-1的数据之和A0" + (i + 1) + "列"
							+ (r1 + j) + "行" + "出错");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构单个账户表格sheet1-1的数据之和A0" + (i + 1) + "列" + (r1 + j)
				+ "行" + "正确");
	}
	/*
	 * 检查1-6每个格子的相加的和
	 */
	public CheckResultMessage checkSum3(int i, int j) {
		int r1 = get(10, 2);
		int c1 = get(11, 2);
		int tr1 = get(20, 5);
		int tc1 = get(21, 5);
		BigDecimal sum = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			for (File f : Main1.files1) {
				try {
					in = new FileInputStream(f);
					hWorkbook = new HSSFWorkbook(in);
					sum = sum.add(getValue(r1 + i, c1 + j,2));
					in.close();
				} catch (Exception e) {
				}
				
			}
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(tr1 + j, tc1 + i, 5).compareTo(sum)) {
					return error("支付机构单个账户表格sheet1-6的数据之和"+(c1 + i) + "行"
							+ (r1 + j) + "列" + "出错");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			for (File f : Main1.files1) {
				try {
					in = new FileInputStream(f);
					xWorkbook = new XSSFWorkbook(in);
					sum = sum.add(getValue1(r1 + i, c1 + j, 2));
					in.close();
				} catch (Exception e) {
				}
			}
			
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue1(tr1 + i, tc1 + j, 5).compareTo(sum)) {
					return error("支付机构单个账户表格sheet1-6的数据之和" +(c1 + i)+ "行"
							+ (r1 + j) + "列" + "出错");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构单个账户表格sheet1-6的数据之和" +(c1 + i)+ "行"
				+ (r1 + j) + "列" + "正确");
	}
	/*
	 * 检查1-9每个格子的相加的和
	 */
	public CheckResultMessage checkSum4(int i, int j) {
		int r1 = get(17, 2);
		int c1 = get(18, 2);
		int tr1 = get(33, 5);
		int tc1 = get(34, 5);
		BigDecimal sum = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			for (File f : Main1.files1) {
				try {
					in = new FileInputStream(f);
					hWorkbook = new HSSFWorkbook(in);
					sum = sum.add(getValue(r1 + i, c1 + j, 3));
					in.close();
				} catch (Exception e) {
				}
				
			}
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(tr1 + j, tc1 + i, 8).compareTo(sum)) {
					return error("支付机构单个账户表格sheet1-9的数据之和"+(c1 + i) + "行"
							+ (r1 + j) + "列" + "出错");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			for (File f : Main1.files1) {
				try {
					in = new FileInputStream(f);
					xWorkbook = new XSSFWorkbook(in);
					sum = sum.add(getValue1(r1 + i, c1 + j, 3));
					in.close();
				} catch (Exception e) {
				}
			}
			
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue1(tr1 + i, tc1 + j, 8).compareTo(sum)) {
					return error("支付机构单个账户表格sheet1-9的数据之和"+(c1 + i) + "行"
							+ (r1 + j) + "列" + "出错");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构单个账户表格sheet1-9的数据之和" +(c1 + i)+ "行"
				+ (r1 + j) + "列" + "正确");
	}
	/*
	 * 检查1-10每个格子的相加的和
	 */
	public CheckResultMessage checkSum5(int i, int j) {
		int r1 = get(20, 2);
		int c1 = get(21, 2);
		int tr1 = get(36, 5);
		int tc1 = get(37, 5);
		BigDecimal sum = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			for (File f : Main1.files1) {
				try {
					in = new FileInputStream(f);
					hWorkbook = new HSSFWorkbook(in);
					sum = sum.add(getValue(r1 + j, c1 + i, 4));
				} catch (Exception e) {
				}
			}
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(tr1 + j, tc1 + i, 9).compareTo(sum)) {
					return error("支付机构单个账户表格sheet1-10的数据之和" +(r1 + j)+ "行"
							+ (c1 + i) + "列" + "出错");
				}
			} catch (Exception e) {
			}
		} else {
			for (File f : Main1.files1) {
				try {
					in = new FileInputStream(f);
					xWorkbook = new XSSFWorkbook(in);
					sum = sum.add(getValue1(r1 + j, c1 + i, 3));
					
				} catch (Exception e) {
				}
			}
			
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue1(tr1 + j, tc1 + i, 9).compareTo(sum)) {
					return error("支付机构单个账户表格sheet1-10的数据之和"+(r1 + j) + "行"
							+ (c1 + i) + "列" + "出错");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构单个账户表格sheet1-10的数据之和"+(r1 + j) + "行"
				+ (c1 + i) + "列" + "正确");
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
			hSheet = hWorkbook.getSheetAt(sheetat);
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
			hSheet = hWorkbook.getSheetAt(sheetat);
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

	public void date() {
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
			// return error(fileName + " 期数格式不正确！应为yyyyMM");
		}
	}

	public CheckResultMessage pass(String message) {
		return new CheckResultMessage(message);
	}

	public CheckResultMessage error(String message) {
		return new CheckResultMessage(message, CheckResultMessage.CHECK_ERROR);
	}
}
