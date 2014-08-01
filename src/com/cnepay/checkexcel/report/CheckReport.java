package com.cnepay.checkexcel.report;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CheckReport {

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

	public CheckReport(File file) throws Exception {
		this.file = file;

		fileName = file.getName();
		version = checkVersion(file);
	}

	// 单个支付机构勾稽关系

	/*
	 * 增加银行余额的特殊业务合计 F = F01 + F02 + … + F09
	 */
	public CheckResultMessage checkF(int day) {
		BigDecimal f = new BigDecimal(0);
		int row = get(10, 2);
		int clums = get(11, 2);
		int total = get(13, 2);
		if (version.equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				for (int i = row; i < total; i++) {
					f = f.add(getValue(row, clums + day, 2));
				}
				if (0 != (f.compareTo(getValue(total, clums + day, 2)))) {
					return error("支付机构单个账户报表<" + fileName + ">增加银行余额的特殊业务合计 F:"
							+ day + "日错误");
				}
				in.close();
			} catch (IOException e) {
				e.printStackTrace();
			}

		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				for (int i = row; i < total; i++) {
					f = f.add(getValue(row, clums, 2));
				}
				if (0 != (f.compareTo(getValue(total, clums + day, 2)))) {
					return error("支付机构单个账户报表<" + fileName + ">增加银行余额的特殊业务合计 F:"
							+ day + "日错误");
				}
				in.close();
			} catch (Exception e) {

			}
		}
		return pass("支付机构单个账户报表<" + fileName + ">增加银行余额的特殊业务合计 F:" + day
				+ "日正确");
	}

	/*
	 * 减少银行余额的特殊业务合计 G = G01 + G02 + … + G13
	 */
	public CheckResultMessage checkG(int day) {
		BigDecimal g = new BigDecimal(0);

		int row1 = get(10, 2);
		int clums1 = get(11, 2);
		int total1 = get(13, 2);

		if (version.equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				for (int i = row1; i < total1; i++) {
					g = g.add(getValue(row1, clums1 + day, 2));
				}
				if (0 != (g.compareTo(getValue(total1, clums1 + day, 2)))) {
					return error("支付机构单个账户报表<" + fileName + ">增加银行余额的特殊业务合计 G:"
							+ day + "日错误");
				}
				in.close();
			} catch (IOException e) {
				e.printStackTrace();
			}

		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				for (int i = row1; i < total1; i++) {
					g = g.add(getValue(row1, clums1, 2));
				}
				if (0 != (g.compareTo(getValue(total1, clums1 + day, 2)))) {
					return error("支付机构单个账户报表<" + fileName + ">增加银行余额的特殊业务合计 F:"
							+ day + "日错误");
				}
				in.close();
			} catch (Exception e) {

			}
		}
		return pass("支付机构单个账户报表<" + fileName + ">增加银行余额的特殊业务合计 G:" + day
				+ "日正确");
	}

	/*
	 * 头寸调拨 F02 + G01 = 0
	 */
	public CheckResultMessage checkF02(int day) {
		int r1 = get(10, 2);
		int c1 = get(11, 2);
		int r2 = get(13, 2);
		int c2 = get(14, 2);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(r1 + 1, c1 + day, 2).abs().compareTo(
						getValue(r2, c2 + day, 2).abs())) {
					return error("支付机构单个账户报表<" + fileName + ">头寸调拨:" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue(r1 + 1, c1 + day, 2).abs().compareTo(
						getValue(r2, c2 + day, 2).abs())) {
					return error("支付机构单个账户报表<" + fileName + ">头寸调拨:" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {
			}

		}
		return pass("支付机构单个账户报表<" + fileName + ">头寸调拨:" + day + "日正确");
	}

	/*
	 * 利息划转 F07 + G02 = 0
	 */
	public CheckResultMessage checkF07(int day) {
		int r1 = get(10, 2);
		int c1 = get(11, 2);
		int r2 = get(13, 2);
		int c2 = get(14, 2);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(r1 + 6, c1 + day, 2).abs().compareTo(
						getValue(r2 + 1, c2 + day, 2).abs())) {
					return error("支付机构单个账户报表<" + fileName + ">利息划转:" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue(r1 + 6, c1 + day, 2).abs().compareTo(
						getValue(r2 + 1, c2 + day, 2).abs())) {
					return error("支付机构单个账户报表<" + fileName + ">利息划转:" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构单个账户报表<" + fileName + ">利息划转:" + day + "日正确");
	}

	/*
	 * 非活期转活期 F04 + G04 = 0
	 */
	public CheckResultMessage checkF04(int day) {
		int r1 = get(10, 2);
		int c1 = get(11, 2);
		int r2 = get(13, 2);
		int c2 = get(14, 2);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(r1 + 3, c1 + day, 2).abs().compareTo(
						getValue(r2 + 3, c2 + day, 2).abs())) {
					return error("支付机构单个账户报表<" + fileName + ">非活期转活期 :" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue(r1 + 3, c1 + day, 2).abs().compareTo(
						getValue(r2 + 3, c2 + day, 2).abs())) {
					return error("支付机构单个账户报表<" + fileName + ">非活期转活期 :" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构单个账户报表<" + fileName + ">非活期转活期 :" + day + "日正确");
	}

	/*
	 * 活期转非活期 F05 + G05 = 0
	 */
	public CheckResultMessage checkF05(int day) {
		int r1 = get(10, 2);
		int c1 = get(11, 2);
		int r2 = get(13, 2);
		int c2 = get(14, 2);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(r1 + 4, c1 + day, 2).abs().compareTo(
						getValue(r2 + 4, c2 + day, 2).abs())) {
					return error("支付机构单个账户报表<" + fileName + ">活期转非活期:" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue(r1 + 4, c1 + day, 2).abs().compareTo(
						getValue(r2 + 4, c2 + day, 2).abs())) {
					return error("支付机构单个账户报表<" + fileName + ">活期转非活期:" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构单个账户报表<" + fileName + ">活期转非活期:" + day + "日正确");
	}

	/*
	 * 系统已增加银行未增加未达账项余额 J01 = K02 + K04 + K06
	 */
	public CheckResultMessage checkJ01(int day) {
		int r1 = get(17, 2);
		int c1 = get(18, 2);
		int r2 = get(20, 2);
		int c2 = get(21, 2);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r2 + day, c2 + 1, 4).add(
						getValue(r2 + day, c2 + 3, 4)).add(
						getValue(r2 + day, c2 + 5, 4));
				if (0 != getValue(r1, c1 + 1 + day, 3).compareTo(b)) {
					return error("支付机构单个账户报表<" + fileName
							+ ">系统已增加银行未增加未达账项余额 J01:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue1(r2 + day, c2 + 1, 4).add(
						getValue1(r2 + day, c2 + 3, 4)).add(
						getValue1(r2 + day, c2 + 5, 4));
				if (0 != getValue1(r1, c1 + 1 + day, 3).compareTo(b)) {
					return error("支付机构单个账户报表<" + fileName
							+ ">系统已增加银行未增加未达账项余额 J01:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构单个账户报表<" + fileName + ">系统已增加银行未增加未达账项余额 J01:" + day
				+ "日正确");
	}

	/*
	 * 系统已减少银行未减少未达账项余额 J02 = K08 + K10 + K12
	 */
	public CheckResultMessage checkJ02(int day) {
		int r1 = get(17, 2);
		int c1 = get(18, 2);
		int r2 = get(20, 2);
		int c2 = get(21, 2);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r2 + day, c2 + 7, 4).add(
						getValue(r2 + day, c2 + 9, 4)).add(
						getValue(r2 + day, c2 + 11, 4));
				if (0 != getValue(r1 + 1, c1 + 1 + day, 3).compareTo(b)) {
					return error("支付机构单个账户报表<" + fileName
							+ ">系统已减少银行未减少未达账项余额 J02:" + day + "日正确");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue1(r2 + day, c2 + 7, 4).add(
						getValue1(r2 + day, c2 + 9, 4)).add(
						getValue1(r2 + day, c2 + 11, 4));
				if (0 != getValue1(r1 + 1, c1 + 1 + day, 3).compareTo(b)) {
					return error("支付机构单个账户报表<" + fileName
							+ ">系统已减少银行未减少未达账项余额 J02:" + day + "日正确");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构单个账户报表<" + fileName + ">系统已减少银行未减少未达账项余额 J02:" + day
				+ "日正确");
	}

	/*
	 * 系统未增加银行已增加未达账项余额 J03 = K14 + K16 + K18
	 */
	public CheckResultMessage checkJ03(int day) {
		int r1 = get(17, 2);
		int c1 = get(18, 2);
		int r2 = get(20, 2);
		int c2 = get(21, 2);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r2 + day, c2 + 13, 4).add(
						getValue(r2 + day, c2 + 15, 4)).add(
						getValue(r2 + day, c2 + 17, 4));
				if (0 != getValue(r1 + 2, c1 + 1 + day, 3).compareTo(b)) {
					return error("支付机构单个账户报表<" + fileName
							+ ">系统未增加银行已增加未达账项余额 J03:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue1(r2 + day, c2 + 13, 4).add(
						getValue1(r2 + day, c2 + 15, 4)).add(
						getValue1(r2 + day, c2 + 17, 4));
				if (0 != getValue1(r1 + 2, c1 + 1 + day, 3).compareTo(b)) {
					return error("支付机构单个账户报表<" + fileName
							+ ">系统未增加银行已增加未达账项余额 J03:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构单个账户报表<" + fileName + ">系统未增加银行已增加未达账项余额 J03:" + day
				+ "日正确");
	}

	/*
	 * 系统未减少银行已减少未达账项余额 J04 = K20 + K22 + K24
	 */
	public CheckResultMessage checkJ04(int day) {
		int r1 = get(17, 2);
		int c1 = get(18, 2);
		int r2 = get(20, 2);
		int c2 = get(21, 2);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r2 + day, c2 + 19, 4).add(
						getValue(r2 + day, c2 + 21, 4)).add(
						getValue(r2 + day, c2 + 23, 4));
				if (0 != getValue(r1 + 3, c1 + 1 + day, 3).compareTo(b)) {
					return error("支付机构单个账户报表<" + fileName
							+ ">系统未减少银行已减少未达账项余额 J04:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue1(r2 + day, c2 + 19, 4).add(
						getValue1(r2 + day, c2 + 21, 4)).add(
						getValue1(r2 + day, c2 + 23, 4));
				if (0 != getValue1(r1 + 3, c1 + 1 + day, 3).compareTo(b)) {
					return error("支付机构单个账户报表<" + fileName
							+ ">系统未减少银行已减少未达账项余额 J04:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构单个账户报表<" + fileName + ">系统未减少银行已减少未达账项余额 J04:" + day
				+ "日正确");
	}

	// 汇总表勾稽关系

	/*
	 * 银行账户入金 X1 =（A01 + A02 + AO3 + A04 + A05 + A06 + A10）+ B09 + F 银行账户出金
	 * X2=（B04 + B05 + B06）+ A14 - G 出入金核对 X1 – X2 = L24 = M1(T) – M1(T-1)
	 */
	public CheckResultMessage check1(int day) {
		int r1 = get(2, 5);
		int c1 = get(3, 5);
		int r2 = get(6, 5);
		int c2 = get(7, 5);
		int r3 = get(22, 5);
		int c3 = get(21, 5);
		int r4 = get(25, 2);
		int r5 = get(43, 5);
		int c5 = get(44, 5);
		BigDecimal x1 = new BigDecimal(0);
		BigDecimal x2 = new BigDecimal(0);
		BigDecimal m = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				x1 = getValue(r1 + 1 + day, c1, 0)
						.add(getValue(r1 + 1 + day, c1 + 1, 0))
						.add(getValue(r1 + 1 + day, c1 + 2, 0))
						.add(getValue(r1 + 1 + day, c1 + 3, 0))
						.add(getValue(r1 + 1 + day, c1 + 4, 0))
						.add(getValue(r1 + 1 + day, c1 + 5, 0))
						.add(getValue(r1 + 1 + day, c1 + 9, 0))
						.add(getValue(r2 + day, c2 + 8, 1))
						.add(getValue(r3, c1 + day, 5));
				x2 = getValue(r2 + day, c2 + 3, 1)
						.add(getValue(r2 + day, c2 + 4, 1))
						.add(getValue(r2 + day, c2 + 5, 1))
						.add(getValue(r1 + 1 + day, c1 + 13, 0))
						.subtract(getValue(r4, c3 + day, 5));
				m = getValue(r5, c5 + day, 11).subtract(
						getValue(r5, c5 + day - 1, 11));
				if (0 != m.compareTo(x1.subtract(x2))) {
					return error("支付机构汇总报表<" + fileName + ">出入金核对:" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				x1 = getValue1(r1 + 1 + day, c1, 0)
						.add(getValue1(r1 + 1 + day, c1 + 1, 0))
						.add(getValue1(r1 + 1 + day, c1 + 2, 0))
						.add(getValue1(r1 + 1 + day, c1 + 3, 0))
						.add(getValue1(r1 + 1 + day, c1 + 4, 0))
						.add(getValue1(r1 + 1 + day, c1 + 5, 0))
						.add(getValue1(r1 + 1 + day, c1 + 9, 0))
						.add(getValue1(r2 + day, c2 + 8, 1))
						.add(getValue1(r3, c1 + day, 5));
				x2 = getValue1(r2 + day, c2 + 3, 1)
						.add(getValue1(r2 + day, c2 + 4, 1))
						.add(getValue1(r2 + day, c2 + 5, 1))
						.add(getValue1(r1 + 1 + day, c1 + 13, 0))
						.subtract(getValue1(r4, c3 + day, 5));
				m = getValue1(r5, c5 + day, 11).subtract(
						getValue1(r5, c5 + day - 1, 11));
				if (0 != m.compareTo(x1.subtract(x2))) {
					return error("支付机构汇总报表<" + fileName + ">出入金核对:" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">出入金核对:" + day + "日正确");
	}

	/*
	 * 业务系统借记客户资金金额 B01 = B02 + B03
	 */
	public CheckResultMessage check2(int day) {

		BigDecimal a1 = new BigDecimal(0);
		int r1 = get(6, 5);
		int c1 = get(7, 5);
		if (version.equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				a1 = getValue(r1 + day, c1 + 1, 1).add(
						getValue(r1 + day, c1 + 2, 1));
				if (0 != getValue(r1 + day, c1, 1).compareTo(a1)) {
					return error("支付机构汇总报表<" + fileName + ">业务系统借记客户资金金额:"
							+ day + "日错误");

				}
				in.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				a1 = getValue1(r1 + day, c1 + 1, 1).add(
						getValue1(r1 + day, c1 + 2, 1));
				if (0 != getValue1(r1 + day, c1, 1).compareTo(a1)) {
					return error("支付机构汇总报表<" + fileName + ">业务系统借记客户资金金额:"
							+ day + "日错误");
				}
				in.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">业务系统借记客户资金金额:" + day + "日正确");
	}

	/*
	 * 业务出金金额 C = B04 + B05
	 */
	public CheckResultMessage check3(int day) {
		BigDecimal a1 = new BigDecimal(0);
		int r1 = get(6, 5);
		int c1 = get(7, 5);
		int r2 = get(12, 5);
		int c2 = get(11, 5);
		if (version.equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				a1 = getValue(r1 + day, c1 + 3, 1).add(
						getValue(r1 + day, c1 + 5, 1));
				if (0 != getValue(r2, c2 + day, 2).compareTo(a1)) {
					return error("支付机构汇总报表<" + fileName + ">业务出金金额:" + day
							+ "日错误");
				}
				in.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				a1 = getValue1(r1 + day, c1 + 3, 1).add(
						getValue1(r1 + day, c1 + 5, 1));
				if (0 != getValue1(r2, c2 + day, 2).compareTo(a1)) {
					return error("支付机构汇总报表<" + fileName + ">业务出金金额:" + day
							+ "日错误");
				}
				in.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">业务出金金额:" + day + "日正确");
	}

	/*
	 * 客户账户内部转账手续费收入 D03 = D02 – D01
	 */
	public CheckResultMessage check4(int day) {
		BigDecimal a1 = new BigDecimal(0);
		int r3 = get(13, 5);
		int c3 = get(14, 5);
		if (version.equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				a1 = getValue(r3 + day, c3 + 1, 3).subtract(
						getValue(r3 + day, c3, 3));
				if (0 != getValue(r3 + day, c3 + 2, 3).compareTo(a1)) {
					return error("支付机构汇总报表<" + fileName + ">客户账户内部转账手续费收入:"
							+ day + "日错误");
				}
				in.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				a1 = getValue1(r3 + day, c3 + 1, 3).add(
						getValue1(r3 + day, c3 + 2, 3));
				if (0 != getValue1(r3 + day, c3 + 2, 3).compareTo(a1)) {
					return error("支付机构汇总报表<" + fileName + ">客户账户内部转账手续费收入:"
							+ day + "日错误");
				}
				in.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">客户账户内部转账手续费收入:" + day + "日正确");
	}

	/*
	 * 客户资金账户余额延续性 E01(T) = E06(T-1)
	 */
	public CheckResultMessage check5(int day) {
		int r = get(17, 5);
		int c = get(18, 5);

		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != (getValue(r + day, c, 4).compareTo(getValue(r + day
						- 1, c + 5, 4)))) {
					return error("支付机构汇总报表<" + fileName + ">客户资金账户余额延续性E01:"
							+ day + "日错误");
				}
				in.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != (getValue1(r + day, c, 4).compareTo(getValue1(r + day
						- 1, c + 5, 4)))) {
					return error("支付机构汇总报表<" + fileName + ">客户资金账户余额延续性:" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">客户资金账户余额延续性:" + day + "日正确");
	}

	/*
	 * 入金贷记客户资金账户金额 E02 = A01 + A07 + A11 + H02
	 */
	public CheckResultMessage check6(int day) {
		int r1 = get(2, 5);
		int c1 = get(3, 5);
		int r2 = get(17, 5);
		int c2 = get(18, 5);
		int r3 = get(27, 5);
		int c3 = get(28, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r1 + 1 + day, c1, 0)
						.add(getValue(r1 + 1 + day, c1 + 6, 0))
						.add(getValue(r1 + 1 + day, c1 + 10, 0))
						.add(getValue(r3 + 2 + day, c3 + 1, 6));
				if (0 != getValue(r2 + day, c2 + 1, 4).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">入金贷记客户资金账户金额:"
							+ day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r1 + 1 + day, c1, 0)
						.add(getValue1(r1 + 1 + day, c1 + 6, 0))
						.add(getValue1(r1 + 1 + day, c1 + 10, 0))
						.add(getValue1(r3 + 2 + day, c3 + 1, 6));
				if (0 != getValue1(r2 + day, c2 + 1, 4).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">入金贷记客户资金账户金额:"
							+ day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">入金贷记客户资金账户金额:" + day + "日正确");
	}

	/*
	 * 出金借记客户资金账户金额 E03 = B01 + B07 + I02
	 */
	public CheckResultMessage check7(int day) {
		int r1 = get(6, 5);
		int c1 = get(7, 5);
		int r2 = get(17, 5);
		int c2 = get(18, 5);
		int r3 = get(30, 5);
		int c3 = get(31, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r1 + day, c1, 1)
						.add(getValue(r1 + day, c1 + 6, 1)).add(
								getValue(r3 + 2 + day, c3 + 1, 7));
				if (0 != getValue(r2 + day, c2 + 2, 4).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">出金借记客户资金账户金额:"
							+ day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r1 + day, c1, 1).add(
						getValue1(r1 + day, c1 + 6, 1)).add(
						getValue1(r3 + 2 + day, c3 + 1, 7));
				if (0 != getValue1(r2 + day, c2 + 2, 4).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">出金借记客户资金账户金额:"
							+ day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">出金借记客户资金账户金额:" + day + "日正确");
	}

	/*
	 * 客户资金账户借方发生额 E04 = D01
	 */
	public CheckResultMessage checkE04(int day) {
		int r1 = get(13, 5);
		int c1 = get(14, 5);
		int r2 = get(17, 5);
		int c2 = get(18, 5);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(r1 + day, c1, 3).compareTo(
						getValue(r2 + day, c2 + 3, 4))) {
					return error("支付机构汇总报表<" + fileName + ">客户资金账户借方发生额 E04:"
							+ day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue1(r1 + day, c1, 3).compareTo(
						getValue1(r2 + day, c2 + 3, 4))) {
					return error("支付机构汇总报表<" + fileName + ">客户资金账户借方发生额 E04:"
							+ day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">客户资金账户借方发生额 E04:" + day + "日正确");
	}

	/*
	 * 客户资金账户贷方发生额 E05 = D02
	 */
	public CheckResultMessage checkE05(int day) {
		int r1 = get(13, 5);
		int c1 = get(14, 5);
		int r2 = get(17, 5);
		int c2 = get(18, 5);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(r1 + day, c1 + 1, 3).compareTo(
						getValue(r2 + day, c2 + 4, 4))) {
					return error("支付机构汇总报表<" + fileName + ">客户资金账户贷方发生额 E05:"
							+ day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue1(r1 + day, c1 + 1, 3).compareTo(
						getValue1(r2 + day, c2 + 4, 4))) {
					return error("支付机构汇总报表<" + fileName + ">客户资金账户贷方发生额 E05:"
							+ day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">客户资金账户贷方发生额 E05:" + day + "日正确");
	}

	/*
	 * 客户资金账户日终余额 E06 = E01 + E02 – E03 + (E05 – E04)
	 */
	public CheckResultMessage check8(int day) {
		int r = get(17, 5);
		int c = get(18, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r + day, c, 4).add(getValue(r + day, c + 1, 4))
						.subtract(getValue(r + day, c + 2, 4))
						.add(getValue(r + day, c + 4, 4))
						.subtract(getValue(r + day, c + 3, 4));
				if (0 != getValue(r + day, c + 5, 4).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">客户资金账户日终余额:" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {

			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r + day, c, 4).add(getValue1(r + day, c + 1, 4))
						.subtract(getValue1(r + day, c + 2, 4))
						.add(getValue1(r + day, c + 4, 4))
						.subtract(getValue1(r + day, c + 3, 4));
				if (0 != getValue1(r + day, c + 5, 4).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">客户资金账户日终余额:" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {

			}
		}
		return error("支付机构汇总报表<" + fileName + ">客户资金账户日终余额:" + day + "日正确");
	}

	/*
	 * 增加银行余额的特殊业务合计 F = F01 + F02 + … + F09
	 */
	public CheckResultMessage checkF1(int day) {
		BigDecimal f = new BigDecimal(0);
		int row = get(20, 5);
		int clums = get(21, 5);
		int total = get(22, 5);

		if (version.equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				for (int i = row; i < total; i++) {
					f = f.add(getValue(row, clums + day, 2));
				}
				if (0 != (f.compareTo(getValue(total, clums + day, 2)))) {
					return error("支付机构汇总报表<" + fileName + ">增加银行余额的特殊业务合计 F:"
							+ day + "日错误");
				}
				in.close();
			} catch (IOException e) {
				e.printStackTrace();
			}

		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				for (int i = row; i < total; i++) {
					f = f.add(getValue(row, clums, 2));
				}
				if (0 != (f.compareTo(getValue(total, clums + day, 2)))) {
					return error("支付机构汇总报表<" + fileName + ">增加银行余额的特殊业务合计 F:"
							+ day + "日错误");
				}
				in.close();
			} catch (Exception e) {

			}
		}
		return pass("支付机构汇总报表<" + fileName + ">增加银行余额的特殊业务合计 F:" + day + "日正确");
	}

	/*
	 * 减少银行余额的特殊业务合计 G = G01 + G02 + … + G13
	 */
	public CheckResultMessage checkG1(int day) {
		BigDecimal g = new BigDecimal(0);

		int row1 = get(23, 5);
		int clums1 = get(24, 5);
		int total1 = get(25, 5);

		if (version.equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				for (int i = row1; i < total1; i++) {
					g = g.add(getValue(row1, clums1 + day, 2));
				}
				if (0 != (g.compareTo(getValue(total1, clums1 + day, 2)))) {
					return error("支付机构汇总报表<" + fileName + ">增加银行余额的特殊业务合计 G:"
							+ day + "日错误");
				}
				in.close();
			} catch (IOException e) {
				e.printStackTrace();
			}

		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				for (int i = row1; i < total1; i++) {
					g = g.add(getValue(row1, clums1, 2));
				}
				if (0 != (g.compareTo(getValue(total1, clums1 + day, 2)))) {
					return error("支付机构汇总报表<" + fileName + ">增加银行余额的特殊业务合计 F:"
							+ day + "日错误");
				}
				in.close();
			} catch (Exception e) {

			}
		}
		return pass("支付机构汇总报表<" + fileName + ">增加银行余额的特殊业务合计 G:" + day + "日正确");
	}

	/*
	 * 向备付金银行缴存现金形式备付金 F01 = H03
	 */
	public CheckResultMessage checkF01(int day) {
		int r1 = get(20, 5);
		int c1 = get(21, 5);
		int r2 = get(27, 5);
		int c2 = get(28, 5);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(r1, c1 + day, 5).compareTo(
						getValue(r2 + 2 + day, c2 + 2, 6))) {
					return error("支付机构汇总报表<" + fileName
							+ ">向备付金银行缴存现金形式备付金 F01 = H03:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue1(r1, c1 + day, 5).compareTo(
						getValue1(r2 + 2 + day, c2 + 2, 6))) {
					return error("支付机构汇总报表<" + fileName
							+ ">向备付金银行缴存现金形式备付金 F01 = H03:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">向备付金银行缴存现金形式备付金 F01 = H03:"
				+ day + "日正确");
	}

	/*
	 * 向备付金银行缴存现金形式预付卡押金F08 = N3
	 */
	public CheckResultMessage checkF08(int day) {
		int r1 = get(20, 5);
		int c1 = get(21, 5);
		int r2 = get(47, 5);
		int c2 = get(48, 5);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(r1 + 8, c1 + day, 5).compareTo(
						getValue(r2 + day, c2 + 2, 12))) {
					return error("支付机构汇总报表<" + fileName
							+ ">向备付金银行缴存现金形式预付卡押金F08 = N3:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue(r1 + 8, c1 + day, 5).compareTo(
						getValue(r2 + day, c2 + 2, 12))) {
					return error("支付机构汇总报表<" + fileName
							+ ">向备付金银行缴存现金形式预付卡押金F08 = N3:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">向备付金银行缴存现金形式预付卡押金F08 = N3:"
				+ day + "日正确");
	}

	/*
	 * 办理预付卡先行现金赎回业务 G09 + I03 = 0
	 */
	public CheckResultMessage checkG09(int day) {
		int r1 = get(23, 5);
		int c1 = get(24, 5);
		int r2 = get(30, 5);
		int c2 = get(31, 5);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != (getValue(r1 + 7, c1 + day, 5).add(
						getValue(r2 + 2 + day, c2 + 2, 7))
						.compareTo(new BigDecimal(0)))) {
					return error("支付机构汇总报表<" + fileName
							+ ">办理预付卡先行现金赎回业务 G09 + I03 = 0:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != (getValue1(r1 + 7, c1 + day, 5).add(
						getValue1(r2 + 2 + day, c2 + 2, 7))
						.compareTo(new BigDecimal(0)))) {
					return error("支付机构汇总报表<" + fileName
							+ ">办理预付卡先行现金赎回业务 G09 + I03 = 0:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">向备付金银行缴存现金形式预付卡押金F08 = N3:"
				+ day + "日正确");
	}

	/*
	 * 以转账方式退回购卡押金 G11 + N4 = 0
	 */
	public CheckResultMessage checkG11(int day) {
		int r1 = get(23, 5);
		int c1 = get(24, 5);
		int r2 = get(47, 5);
		int c2 = get(48, 5);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(r1 + 11, c1 + day, 5).compareTo(
						getValue(r2 + day, c2 + 3, 12))) {
					return error("支付机构汇总报表<" + fileName
							+ ">以转账方式退回购卡押金 G11 + N4 = 0:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue(r1 + 11, c1 + day, 5).compareTo(
						getValue(r2 + day, c2 + 3, 12))) {
					return error("支付机构汇总报表<" + fileName
							+ ">以转账方式退回购卡押金 G11+N4=0:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">以转账方式退回购卡押金 G11+N4=0:" + day
				+ "日正确");
	}

	/*
	 * 办理预付卡押金先行现金赎回业务 G12 + N5 = 0
	 */
	public CheckResultMessage checkG12(int day) {
		int r1 = get(23, 5);
		int c1 = get(24, 5);
		int r2 = get(47, 5);
		int c2 = get(48, 5);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(r1 + 12, c1 + day, 5).compareTo(
						getValue(r2 + day, c2 + 4, 12))) {
					return error("支付机构汇总报表<" + fileName
							+ ">办理预付卡押金先行现金赎回业务 G12 + N5 = 0:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue(r1 + 12, c1 + day, 5).compareTo(
						getValue(r2 + day, c2 + 4, 12))) {
					return error("支付机构汇总报表<" + fileName
							+ ">办理预付卡押金先行现金赎回业务 G12 + N5 = 0:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">办理预付卡押金先行现金赎回业务 G12 + N5 = 0:"
				+ day + "日正确");
	}
	/*
	 * 头寸调拨 F02 + G01 = 0
	 */
	public CheckResultMessage checkTF02(int day) {
		int r1 = get(20, 2);
		int c1 = get(21, 2);
		int r2 = get(23, 2);
		int c2 = get(24, 2);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(r1 + 1, c1 + day, 2).abs().compareTo(
						getValue(r2, c2 + day, 2).abs())) {
					return error("支付机构单个账户报表<" + fileName + ">头寸调拨:" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue(r1 + 1, c1 + day, 2).abs().compareTo(
						getValue(r2, c2 + day, 2).abs())) {
					return error("支付机构单个账户报表<" + fileName + ">头寸调拨:" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {
			}

		}
		return pass("支付机构单个账户报表<" + fileName + ">头寸调拨:" + day + "日正确");
	}

	/*
	 * 利息划转 F07 + G02 = 0
	 */
	public CheckResultMessage checkTF07(int day) {
		int r1 = get(20, 2);
		int c1 = get(21, 2);
		int r2 = get(23, 2);
		int c2 = get(24, 2);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(r1 + 6, c1 + day, 2).abs().compareTo(
						getValue(r2 + 1, c2 + day, 2).abs())) {
					return error("支付机构单个账户报表<" + fileName + ">利息划转:" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue(r1 + 6, c1 + day, 2).abs().compareTo(
						getValue(r2 + 1, c2 + day, 2).abs())) {
					return error("支付机构单个账户报表<" + fileName + ">利息划转:" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构单个账户报表<" + fileName + ">利息划转:" + day + "日正确");
	}

	/*
	 * 非活期转活期 F04 + G04 = 0
	 */
	public CheckResultMessage checkTF04(int day) {
		int r1 = get(20, 2);
		int c1 = get(21, 2);
		int r2 = get(23, 2);
		int c2 = get(24, 2);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(r1 + 3, c1 + day, 2).abs().compareTo(
						getValue(r2 + 3, c2 + day, 2).abs())) {
					return error("支付机构单个账户报表<" + fileName + ">非活期转活期 :" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue(r1 + 3, c1 + day, 2).abs().compareTo(
						getValue(r2 + 3, c2 + day, 2).abs())) {
					return error("支付机构单个账户报表<" + fileName + ">非活期转活期 :" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构单个账户报表<" + fileName + ">非活期转活期 :" + day + "日正确");
	}

	/*
	 * 活期转非活期 F05 + G05 = 0
	 */
	public CheckResultMessage checkTF05(int day) {
		int r1 = get(20, 2);
		int c1 = get(21, 2);
		int r2 = get(23, 2);
		int c2 = get(24, 2);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(r1 + 4, c1 + day, 2).abs().compareTo(
						getValue(r2 + 4, c2 + day, 2).abs())) {
					return error("支付机构单个账户报表<" + fileName + ">活期转非活期:" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue(r1 + 4, c1 + day, 2).abs().compareTo(
						getValue(r2 + 4, c2 + day, 2).abs())) {
					return error("支付机构单个账户报表<" + fileName + ">活期转非活期:" + day
							+ "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构单个账户报表<" + fileName + ">活期转非活期:" + day + "日正确");
	}
	/*
	 * 现金形式缴存备付金余额延续性 H01(T) = H04(T-1)
	 */
	public CheckResultMessage check9(int day) {
		int r = get(27, 5);
		int c = get(28, 5);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(r + 3 + day, c, 6).compareTo(
						getValue(r + 2 + day, c + 3, 6))) {
					return error("支付机构汇总报表<" + fileName + ">现金形式缴存备付金余额延续性:"
							+ day + "日错误");
				}
			} catch (Exception e) {

			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue1(r + 3 + day, c, 6).compareTo(
						getValue1(r + 2 + day, c + 3, 6))) {
					return error("支付机构汇总报表<" + fileName + ">现金形式缴存备付金余额延续性:"
							+ day + "日错误");
				}
			} catch (Exception e) {

			}
		}
		return pass("支付机构汇总报表<" + fileName + ">现金形式缴存备付金余额延续性:" + day + "日正确");
	}

	/*
	 * 变动额核对 H01 + H02 – H03 = H04
	 */
	public CheckResultMessage check10(int day) {
		int r = get(27, 5);
		int c = get(28, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r + 3 + day, c, 6).add(
						getValue(r + 3 + day, c + 1, 6)).subtract(
						getValue(r + 3 + day, c + 2, 6));
				if (0 != getValue(r + 3 + day, c + 3, 6).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">变动额核对H:" + day
							+ "日错误");
				}
			} catch (Exception e) {

			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r + 3 + day, c, 6).add(
						getValue1(r + 3 + day, c + 1, 6)).subtract(
						getValue1(r + 3 + day, c + 2, 6));
				if (0 != getValue1(r + 3 + day, c + 3, 6).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">变动额核对H:" + day
							+ "日错误");
				}
			} catch (Exception e) {

			}
		}
		return pass("支付机构汇总报表<" + fileName + ">变动额核对H:" + day + "日正确");
	}

	/*
	 * 以自有资金赎回预付卡余额延续性I01(T) = I04(T-1)
	 */
	public CheckResultMessage check11(int day) {
		int r = get(30, 5);
		int c = get(31, 5);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(r + 3 + day, c, 7).compareTo(
						getValue(r + 2 + day, c + 3, 7))) {
					return error("支付机构汇总报表<" + fileName + ">以自有资金赎回预付卡余额延续性:"
							+ day + "日错误");
				}
			} catch (Exception e) {

			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue1(r + 3 + day, c, 7).compareTo(
						getValue1(r + 2 + day, c + 3, 7))) {
					return error("支付机构汇总报表<" + fileName + ">以自有资金赎回预付卡余额延续性:"
							+ day + "日错误");
				}
			} catch (Exception e) {

			}
		}
		return pass("支付机构汇总报表<" + fileName + ">以自有资金赎回预付卡余额延续性:" + day + "日正确");
	}

	/*
	 * 变动额核对 I01 + I02 – I03 = I04
	 */
	public CheckResultMessage check12(int day) {
		int r = get(30, 5);
		int c = get(31, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r + 3 + day, c, 7).add(
						getValue(r + 3 + day, c + 1, 7)).subtract(
						getValue(r + 3 + day, c + 2, 7));
				if (0 != getValue(r + 3 + day, c + 3, 7).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">变动额核对I:" + day
							+ "日错误");
				}
			} catch (Exception e) {

			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r + 3 + day, c, 7).add(
						getValue1(r + 3 + day, c + 1, 7)).subtract(
						getValue1(r + 3 + day, c + 2, 7));
				if (0 != getValue1(r + 3 + day, c + 3, 7).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">变动额核对I:" + day
							+ "日错误");
				}
			} catch (Exception e) {

			}
		}
		return pass("支付机构汇总报表<" + fileName + ">变动额核对I:" + day + "日正确");
	}

	/*
	 * 系统已增加银行未增加未达账项余额 J01 = K02 + K04 + K06
	 */
	public CheckResultMessage check13(int day) {
		int r1 = get(33, 5);
		int c1 = get(34, 5);
		int r2 = get(36, 5);
		int c2 = get(37, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r2 + 1 + day, c2 + 1, 9).add(
						getValue(r2 + 1 + day, c2 + 3, 9)).add(
						getValue(r2 + 1 + day, c2 + 5, 9));
				if (0 != getValue(r1, c1 + 1 + day, 8).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">系统已增加银行未增加未达账项余额:"
							+ day + "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r2 + 1 + day, c2 + 1, 9).add(
						getValue1(r2 + 1 + day, c2 + 3, 9)).add(
						getValue1(r2 + 1 + day, c2 + 5, 9));
				if (0 != getValue1(r1, c1 + 1 + day, 8).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">系统已增加银行未增加未达账项余额:"
							+ day + "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">系统已增加银行未增加未达账项余额:" + day + "日正确");
	}

	/*
	 * 系统已减少银行未减少未达账项余额 J02 = K08 + K10 + K12
	 */
	public CheckResultMessage check14(int day) {
		int r1 = get(33, 5);
		int c1 = get(34, 5);
		int r2 = get(36, 5);
		int c2 = get(37, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r2 + 1 + day, c2 + 7, 9).add(
						getValue(r2 + 1 + day, c2 + 9, 9)).add(
						getValue(r2 + 1 + day, c2 + 11, 9));
				if (0 != getValue(r1 + 1, c1 + 1 + day, 8).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">系统已减少银行未减少未达账项余额:"
							+ day + "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r2 + 1 + day, c2 + 7, 9).add(
						getValue1(r2 + 1 + day, c2 + 9, 9)).add(
						getValue1(r2 + 1 + day, c2 + 11, 9));
				if (0 != getValue1(r1, c1 + 1 + day, 8).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">系统已减少银行未减少未达账项余额:"
							+ day + "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">系统已减少银行未减少未达账项余额:" + day + "日正确");
	}

	/*
	 * 系统未增加银行已增加未达账项余额 J03 = K14 + K16 + K18
	 */
	public CheckResultMessage check15(int day) {
		int r1 = get(33, 5);
		int c1 = get(34, 5);
		int r2 = get(36, 5);
		int c2 = get(37, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r2 + 1 + day, c2 + 13, 9).add(
						getValue(r2 + 1 + day, c2 + 15, 9)).add(
						getValue(r2 + 1 + day, c2 + 17, 9));
				if (0 != getValue(r1 + 2, c1 + 1 + day, 8).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">系统未增加银行已增加未达账项余额:"
							+ day + "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r2 + 1 + day, c2 + 13, 9).add(
						getValue1(r2 + 1 + day, c2 + 15, 9)).add(
						getValue1(r2 + 1 + day, c2 + 17, 9));
				if (0 != getValue1(r1 + 2, c1 + 1 + day, 8).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">系统未增加银行已增加未达账项余额:"
							+ day + "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">系统未增加银行已增加未达账项余额:" + day + "日正确");
	}

	/*
	 * 系统未减少银行已减少未达账项余额 J04 = K20 + K22 + K24
	 */
	public CheckResultMessage check16(int day) {
		int r1 = get(33, 5);
		int c1 = get(34, 5);
		int r2 = get(36, 5);
		int c2 = get(37, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r2 + 1 + day, c2 + 19, 9).add(
						getValue(r2 + 1 + day, c2 + 21, 9)).add(
						getValue(r2 + 1 + day, c2 + 23, 9));
				if (0 != getValue(r1 + 3, c1 + 1 + day, 8).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">系统未减少银行已减少未达账项余额:"
							+ day + "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r2 + 1 + day, c2 + 19, 9).add(
						getValue1(r2 + 1 + day, c2 + 21, 9)).add(
						getValue1(r2 + 1 + day, c2 + 23, 9));
				if (0 != getValue1(r1 + 3, c1 + 1 + day, 8).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">系统未减少银行已减少未达账项余额:"
							+ day + "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">系统未减少银行已减少未达账项余额:" + day + "日正确");
	}
	/*
	 * 期初客户资金余额 L1 = E01
	 */
	public CheckResultMessage checkL1(int day){
		int r1 = get(17,5);
		int c1 = get(18,5);
		int r2 = get(39,5);
		int c2 = get(40,5);
		if(checkVersion(file).equals("2003")){
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if(0!=getValue(r2, c2+day, 10).compareTo(getValue(r1+day, c1, 4))){
					return error("支付机构汇总报表<" + fileName + ">期初客户资金余额 L1 = E01:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}else{
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if(0!=getValue1(r2, c2+day, 10).compareTo(getValue1(r1+day, c1, 4))){
					return error("支付机构汇总报表<" + fileName + ">期初客户资金余额 L1 = E01:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">期初客户资金余额 L1 = E01:" + day + "日错误");
	}
	/*
	 * 期末客户资金余额 L2 = E06
	 */
	public CheckResultMessage checkL2(int day){
		int r1 = get(17,5);
		int c1 = get(18,5);
		int r2 = get(39,5);
		int c2 = get(40,5);
		if(checkVersion(file).equals("2003")){
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if(0!=getValue(r2+1, c2+day, 10).compareTo(getValue(r1+day, c1+5, 4))){
					return error("支付机构汇总报表<" + fileName + ">期末客户资金余额 L2 = E06:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}else{
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if(0!=getValue1(r2+1, c2+day, 10).compareTo(getValue1(r1+day, c1+5, 4))){
					return error("支付机构汇总报表<" + fileName + ">期末客户资金余额 L2 = E06:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">期末客户资金余额 L2 = E06:" + day + "日错误");
	}
	/*
	 * 本期接受现金形式的客户备付金金额 L4 = H02
	 */
	public CheckResultMessage checkL4(int day){
		int r1 = get(17,5);
		int c1 = get(18,5);
		int r2 = get(27,5);
		int c2 = get(28,5);
		if(checkVersion(file).equals("2003")){
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if(0!=getValue(r2+4, c2+day, 10).compareTo(getValue(r1+2+day, c1+1, 6))){
					return error("支付机构汇总报表<" + fileName + ">本期接受现金形式的客户备付金金额 L4 = H02:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}else{
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if(0!=getValue(r2+4, c2+day, 10).compareTo(getValue(r1+2+day, c1+1, 6))){
					return error("支付机构汇总报表<" + fileName + ">本期接受现金形式的客户备付金金额 L4 = H02:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">本期接受现金形式的客户备付金金额 L4 = H02:" + day + "日错误");
	}
	/*
	 * 本期向备付金银行缴存现金备付金 L5 = H03 
	 */
	public CheckResultMessage checkL5(int day){
		int r1 = get(17,5);
		int c1 = get(18,5);
		int r2 = get(27,5);
		int c2 = get(28,5);
		if(checkVersion(file).equals("2003")){
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if(0!=getValue(r2+5, c2+day, 10).compareTo(getValue(r1+2+day, c1+2, 6))){
					return error("支付机构汇总报表<" + fileName + ">本期向备付金银行缴存现金备付金 L5 = H03:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}else{
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if(0!=getValue(r2+5, c2+day, 10).compareTo(getValue(r1+2+day, c1+2, 6))){
					return error("支付机构汇总报表<" + fileName + ">本期向备付金银行缴存现金备付金 L5 = H03:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">本期向备付金银行缴存现金备付金 L5 = H03:" + day + "日错误");
	}
	/*
	 * 本期以自有资金先行赎回预付卡的金额 L6 = I02
	 */
	public CheckResultMessage checkL6(int day){
		int r1 = get(17,5);
		int c1 = get(18,5);
		int r2 = get(30,5);
		int c2 = get(31,5);
		if(checkVersion(file).equals("2003")){
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if(0!=getValue(r2+6, c2+day, 10).compareTo(getValue(r1+2+day, c1+1, 6))){
					return error("支付机构汇总报表<" + fileName + ">本期以自有资金先行赎回预付卡的金额 L6 = I02:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}else{
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if(0!=getValue(r2+6, c2+day, 10).compareTo(getValue(r1+2+day, c1+1, 6))){
					return error("支付机构汇总报表<" + fileName + ">本期以自有资金先行赎回预付卡的金额 L6 = I02:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">本期以自有资金先行赎回预付卡的金额 L6 = I02:" + day + "日错误");
	}
	/*
	 * 本期向备付金存管银行办理预付卡先行赎回资金结转业务金额 L7 = I03 
	 */
	public CheckResultMessage checkL7(int day){
		int r1 = get(17,5);
		int c1 = get(18,5);
		int r2 = get(30,5);
		int c2 = get(31,5);
		if(checkVersion(file).equals("2003")){
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if(0!=getValue(r2+7, c2+day, 10).compareTo(getValue(r1+2+day, c1+2, 6))){
					return error("支付机构汇总报表<" + fileName + ">本期向备付金存管银行办理预付卡先行赎回资金结转业务金额 L7 = I03:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}else{
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if(0!=getValue(r2+7, c2+day, 10).compareTo(getValue(r1+2+day, c1+2, 6))){
					return error("支付机构汇总报表<" + fileName + ">本期向备付金存管银行办理预付卡先行赎回资金结转业务金额 L7 = I03:" + day + "日错误");
				}
				in.close();
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">本期向备付金存管银行办理预付卡先行赎回资金结转业务金额 L7 = I03:" + day + "日错误");
	}
	/*
	 * 未达账项核对 J01(T) = J01(T-1) + L9
	 */
	public CheckResultMessage check17(int day) {
		int r1 = get(33, 5);
		int c1 = get(34, 5);
		int r2 = get(39, 5);
		int c2 = get(40, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r1, c1 + day, 8).add(
						getValue(r2 + 9, c2 + day - 1, 10));
				if (0 != getValue(r1, c1 + 1 + day, 8).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">未达账项核对 J01:" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r1, c1 + day, 8).add(
						getValue1(r2 + 9, c2 + day - 1, 10));
				if (0 != getValue1(r1, c1 + 1 + day, 8).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">未达账项核对 J01:" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">未达账项核对 J01:" + day + "日正确");
	}

	/*
	 * J02(T) = J02(T-1) + L10
	 */
	public CheckResultMessage check18(int day) {
		int r1 = get(33, 5);
		int c1 = get(34, 5);
		int r2 = get(39, 5);
		int c2 = get(40, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r1 + 1, c1 + day, 8).add(
						getValue(r2 + 10, c2 + day - 1, 10));
				if (0 != getValue(r1 + 1, c1 + 1 + day, 8).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">未达账项核对 J02:" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r1 + 1, c1 + day, 8).add(
						getValue1(r2 + 10, c2 + day - 1, 10));
				if (0 != getValue1(r1 + 1, c1 + 1 + day, 8).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">未达账项核对 J02:" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">未达账项核对 J02:" + day + "日正确");
	}

	/*
	 * 未达账项核对 J03(T) = J03(T-1) + L11
	 */
	public CheckResultMessage check19(int day) {
		int r1 = get(33, 5);
		int c1 = get(34, 5);
		int r2 = get(39, 5);
		int c2 = get(40, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r1 + 2, c1 + day, 8).add(
						getValue(r2 + 11, c2 + day - 1, 10));
				if (0 != getValue(r1 + 2, c1 + 1 + day, 8).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">未达账项核对 J03:" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r1 + 2, c1 + day, 8).add(
						getValue1(r2 + 11, c2 + day - 1, 10));
				if (0 != getValue1(r1 + 2, c1 + 1 + day, 8).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">未达账项核对 J03:" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">未达账项核对 J03:" + day + "日正确");
	}

	/*
	 * 未达账项核对 J04(T) = J04(T-1) + L12
	 */
	public CheckResultMessage check20(int day) {
		int r1 = get(33, 5);
		int c1 = get(34, 5);
		int r2 = get(39, 5);
		int c2 = get(40, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r1 + 3, c1 + day, 8).add(
						getValue(r2 + 12, c2 + day - 1, 10));
				if (0 != getValue(r1 + 3, c1 + 1 + day, 8).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">未达账项核对 J04:" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r1 + 3, c1 + day, 8).add(
						getValue1(r2 + 12, c2 + day - 1, 10));
				if (0 != getValue1(r1 + 3, c1 + 1 + day, 8).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">未达账项核对 J04:" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">未达账项核对 J04:" + day + "日正确");
	}

	/*
	 * 客户资金余额变动额 L3 = L2 - L1
	 */
	public CheckResultMessage check21(int day) {
		int r = get(39, 5);
		int c = get(40, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r + 1, c + day, 10).subtract(
						getValue(r, c + day, 10));
				if (0 != getValue(r + 2, c + day, 10).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">客户资金余额变动额 :" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue1(r + 1, c + day, 10).subtract(
						getValue1(r, c + day, 10));
				if (0 != getValue1(r + 2, c + day, 10).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">客户资金余额变动额 :" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">客户资金余额变动额 :" + day + "日正确");
	}

	/*
	 * 本期实现的手续费收入 L8 = (A02 + A08 + A12) + （B03 + B08） + D03
	 */
	public CheckResultMessage check22(int day) {
		int r1 = get(2, 5);
		int c1 = get(3, 5);
		int r2 = get(6, 5);
		int c2 = get(7, 5);
		int r3 = get(13, 5);
		int c3 = get(14, 5);
		int r4 = get(39, 5);
		int c4 = get(40, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r1 + 1 + day, c1 + 1, 0)
						.add(getValue(r1 + 1 + day, c1 + 7, 0))
						.add(getValue(r1 + 1 + day, c1 + 11, 0))
						.add(getValue(r2 + day, c2 + 2, 1))
						.add(getValue(r2 + day, c2 + 7, 1))
						.add(getValue(r3 + day, c3 + 2, 3));
				if (0 != getValue(r4 + 8, c4 + day, 10).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">本期实现的手续费收入 :" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r1 + 1 + day, c1 + 1, 0)
						.add(getValue1(r1 + 1 + day, c1 + 7, 0))
						.add(getValue1(r1 + 1 + day, c1 + 11, 0))
						.add(getValue1(r2 + day, c2 + 2, 1))
						.add(getValue1(r2 + day, c2 + 7, 1))
						.add(getValue1(r3 + day, c3 + 2, 3));
				if (0 != getValue1(r4 + 8, c4 + day, 10).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">本期实现的手续费收入 :" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">本期实现的手续费收入 :" + day + "日正确");
	}

	/*
	 * 本期支付机构已增加客户资金余额，备付金银行未增加备付金银行账户余额 L9 =(A07 + A08 + A09) – (A04 + A05 +
	 * A06)
	 */
	public CheckResultMessage check23(int day) {
		int r1 = get(2, 5);
		int c1 = get(3, 5);
		int r2 = get(39, 5);
		int c2 = get(40, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r1 + 1 + day, c1 + 6, 0)
						.add(getValue(r1 + 1 + day, c1 + 7, 0))
						.add(getValue(r1 + 1 + day, c1 + 8, 0))
						.subtract(
								getValue(r1 + 1 + day, c1 + 3, 0).add(
										getValue(r1 + 1 + day, c1 + 4, 0)).add(
										getValue(r1 + 1 + day, c1 + 5, 0)));
				if (0 != getValue(r2 + 9, c2 + day, 10).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName
							+ ">本期支付机构已增加客户资金余额，备付金银行未增加备付金银行账户余额 :" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r1 + 1 + day, c1 + 6, 0)
						.add(getValue1(r1 + 1 + day, c1 + 7, 0))
						.add(getValue1(r1 + 1 + day, c1 + 8, 0))
						.subtract(
								getValue1(r1 + 1 + day, c1 + 3, 0)
										.add(getValue1(r1 + 1 + day, c1 + 4, 0))
										.add(getValue1(r1 + 1 + day, c1 + 5, 0)));
				if (0 != getValue1(r2 + 9, c2 + day, 10).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName
							+ ">本期支付机构已增加客户资金余额，备付金银行未增加备付金银行账户余额 :" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName
				+ ">本期支付机构已增加客户资金余额，备付金银行未增加备付金银行账户余额 :" + day + "日正确");
	}

	/*
	 * 本期支付机构已减少客户资金余额，备付金银行未减少备付金银行余额 L10 = A10 – (A11 + A12 + A13 + A14)
	 */
	public CheckResultMessage check24(int day) {
		int r1 = get(2, 5);
		int c1 = get(3, 5);
		int r2 = get(39, 5);
		int c2 = get(40, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r1 + 1 + day, c1 + 9, 0).subtract(
						getValue(r1 + 1 + day, c1 + 10, 0)
								.add(getValue(r1 + 1 + day, c1 + 11, 0))
								.add(getValue(r1 + 1 + day, c1 + 12, 0))
								.add(getValue(r1 + 1 + day, c1 + 13, 0)));
				if (0 != getValue(r2 + 10, c2 + day, 10).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName
							+ ">本期支付机构已减少客户资金余额，备付金银行未减少备付金银行余额 :" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r1 + 1 + day, c1 + 9, 0).subtract(
						getValue1(r1 + 1 + day, c1 + 10, 0)
								.add(getValue1(r1 + 1 + day, c1 + 11, 0))
								.add(getValue1(r1 + 1 + day, c1 + 12, 0))
								.add(getValue1(r1 + 1 + day, c1 + 13, 0)));
				if (0 != getValue1(r2 + 10, c2 + day, 10).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName
							+ ">本期支付机构已减少客户资金余额，备付金银行未减少备付金银行余额 :" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName
				+ ">本期支付机构已减少客户资金余额，备付金银行未减少备付金银行余额 :" + day + "日正确");
	}

	/*
	 * 本期备付金银行已增加备付金银行账户余额，支付机构未增加客户资金余额 L11 = B02 - B04 - B05
	 */
	public CheckResultMessage check25(int day) {
		int r1 = get(6, 5);
		int c1 = get(7, 5);
		int r2 = get(39, 5);
		int c2 = get(40, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r1 + day, c1 + 1, 1).subtract(
						getValue(r1 + day, c1 + 3, 1)).subtract(
						getValue(r1 + day, c1 + 4, 1));
				if (0 != getValue(r2 + 11, c2 + day, 10).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName
							+ ">本期备付金银行已增加备付金银行账户余额，支付机构未增加客户资金余额 :" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r1 + 1 + day, c1 + 9, 0).subtract(
						getValue1(r1 + 1 + day, c1 + 10, 0)
								.add(getValue1(r1 + 1 + day, c1 + 11, 0))
								.add(getValue1(r1 + 1 + day, c1 + 12, 0))
								.add(getValue1(r1 + 1 + day, c1 + 13, 0)));
				if (0 != getValue1(r2 + 10, c2 + day, 10).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName
							+ ">本期备付金银行已增加备付金银行账户余额，支付机构未增加客户资金余额 :" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName
				+ ">本期备付金银行已增加备付金银行账户余额，支付机构未增加客户资金余额 :" + day + "日正确");
	}

	/*
	 * 本期备付金银行已减少备付金银行账户余额，支付机构未减少客户资金余额 L12 = B06 – (B07 + B08 + B09)
	 */
	public CheckResultMessage check26(int day) {
		int r1 = get(6, 5);
		int c1 = get(7, 5);
		int r2 = get(39, 5);
		int c2 = get(40, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r1 + day, c1 + 5, 1).subtract(
						getValue(r1 + day, c1 + 6, 1).add(
								getValue(r1 + day, c1 + 7, 1)).add(
								getValue(r1 + day, c1 + 8, 1)));
				if (0 != getValue(r2 + 12, c2 + day, 10).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName
							+ ">本期备付金银行已减少备付金银行账户余额，支付机构未减少客户资金余额 :" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r1 + 1 + day, c1 + 9, 0).subtract(
						getValue1(r1 + 1 + day, c1 + 10, 0)
								.add(getValue1(r1 + 1 + day, c1 + 11, 0))
								.add(getValue1(r1 + 1 + day, c1 + 12, 0))
								.add(getValue1(r1 + 1 + day, c1 + 13, 0)));
				if (0 != getValue1(r2 + 10, c2 + day, 10).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName
							+ ">本期备付金银行已减少备付金银行账户余额，支付机构未减少客户资金余额 :" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName
				+ ">本期备付金银行已减少备付金银行账户余额，支付机构未减少客户资金余额 :" + day + "日正确");
	}

	/*
	 * 其他调整项 L20 = Z1 + … + Zn
	 */
	public CheckResultMessage check27(int day) {
		int r1 = get(39, 5);
		int c1 = get(40, 5);
		int r2 = get(41, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				for (int i = r1 + 21; i < r2; i++) {
					b = b.add(getValue(i, c1 + day, 10));
				}
				if (0 != getValue(r1 + 20, c1 + day, 10).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">其他跳整项L20:" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				for (int i = r1 + 21; i < r2; i++) {
					b = b.add(getValue1(i, c1 + day, 10));
				}
				if (0 != getValue1(r1 + 20, c1 + day, 10).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">其他跳整项L20:" + day
							+ "日错误");
				}
			} catch (Exception e) {
			}
		}
		return error("支付机构汇总报表<" + fileName + ">其他跳整项L20:" + day + "日正确");
	}

	/*
	 * 客户资金账户变动额试算值 L21 =
	 * L3-L4+L5+L6-L7+L8-L9+L10+L11-L12+L13+L14-L15-L16-L17-L18-L19+L20
	 */
	public CheckResultMessage check28(int day) {
		int r = get(39, 5);
		int c = get(40, 5);
		int r2 = get(41, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r + 2, c + day, 10)
						.subtract(getValue(r + 4, c + day, 10))
						.add(getValue(r + 5, c + day, 10))
						.add(getValue(r + 6, c + day, 10))
						.subtract(getValue(r + 7, c + day, 10))
						.add(getValue(r + 8, c + day, 10))
						.subtract(getValue(r + 9, c + day, 10))
						.add(getValue(r + 10, c + day, 10))
						.add(getValue(r + 11, c + day, 10))
						.subtract(getValue(r + 12, c + day, 10))
						.add(getValue(r + 13, c + day, 10))
						.add(getValue(r + 14, c + day, 10))
						.subtract(getValue(r + 15, c + day, 10))
						.subtract(getValue(r + 16, c + day, 10))
						.subtract(getValue(r + 17, c + day, 10))
						.subtract(getValue(r + 18, c + day, 10))
						.subtract(getValue(r + 19, c + day, 10))
						.add(getValue(r + 20, c + day, 10));
				if (0 != getValue(r2, c + day, 10).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">客户资金账户变动额试算值 L21:"
							+ day + "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r + 2, c + day, 10)
						.subtract(getValue1(r + 4, c + day, 10))
						.add(getValue1(r + 5, c + day, 10))
						.add(getValue1(r + 6, c + day, 10))
						.subtract(getValue1(r + 7, c + day, 10))
						.add(getValue1(r + 8, c + day, 10))
						.subtract(getValue1(r + 9, c + day, 10))
						.add(getValue1(r + 10, c + day, 10))
						.add(getValue1(r + 11, c + day, 10))
						.subtract(getValue1(r + 12, c + day, 10))
						.add(getValue1(r + 13, c + day, 10))
						.add(getValue1(r + 14, c + day, 10))
						.subtract(getValue1(r + 15, c + day, 10))
						.subtract(getValue1(r + 16, c + day, 10))
						.subtract(getValue1(r + 17, c + day, 10))
						.subtract(getValue1(r + 18, c + day, 10))
						.subtract(getValue1(r + 19, c + day, 10))
						.add(getValue1(r + 20, c + day, 10));
				if (0 != getValue1(r2, c + day, 10).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName + ">客户资金账户变动额试算值 L21:"
							+ day + "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">客户资金账户变动额试算值 L21:" + day + "日正确");
	}

	/*
	 * 备付金银行账户中未结转的备付金银行存款利息余额 M2(T) = M2(T-1) + (L13 - L16 - L17)
	 */
	public CheckResultMessage check29(int day) {
		int r1 = get(39, 5);
		int c1 = get(40, 5);
		int r2 = get(43, 5);
		int c2 = get(44, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r2 + 1, c2 + day - 1, 11).subtract(
						getValue(r1 + 13, c1 + day - 1, 10).subtract(
								getValue(r1 + 16, c1 + day - 1, 10)).subtract(
								getValue(r1 + 17, c1 + day - 1, 10)));
				if (0 != getValue(r2 + 1, c2 + day, 11).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName
							+ ">备付金银行账户中未结转的备付金银行存款利息余额 M2:" + day + "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r2 + 1, c2 + day - 1, 11).add(
						getValue1(r1 + 13, c1 + day - 1, 10).subtract(
								getValue1(r1 + 16, c1 + day - 1, 10)).subtract(
								getValue1(r1 + 17, c1 + day - 1, 10)));
				if (0 != getValue1(r2 + 1, c2 + day, 11).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName
							+ ">备付金银行账户中未结转的备付金银行存款利息余额 M2:" + day + "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">备付金银行账户中未结转的备付金银行存款利息余额 M2:"
				+ day + "日正确");
	}

	/*
	 * 备付金银行账户中累计申请存放的自有资金余额M3(T) = M3(T-1) + (L14 - L19)
	 */
	public CheckResultMessage check30(int day) {
		int r1 = get(39, 5);
		int c1 = get(40, 5);
		int r2 = get(43, 5);
		int c2 = get(44, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r2 + 2, c2 + day - 1, 11).add(
						getValue(r1 + 14, c1 + day - 1, 10).subtract(
								getValue(r1 + 19, c1 + day - 1, 10)));
				if (0 != getValue(r2 + 2, c2 + day, 11).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName
							+ ">备付金银行账户中累计申请存放的自有资金余额M3:" + day + "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r2 + 2, c2 + day - 1, 11).add(
						getValue1(r1 + 14, c1 + day - 1, 10).subtract(
								getValue1(r1 + 19, c1 + day - 1, 10)));
				if (0 != getValue1(r2 + 2, c2 + day, 11).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName
							+ ">备付金银行账户中累计申请存放的自有资金余额M3:" + day + "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">备付金银行账户中累计申请存放的自有资金余额M3:" + day
				+ "日正确");
	}

	/*
	 * 备付金银行账户中未结转的支付业务净收入余额 M4(T) = M4(T-1) + (L8 - L15 - L18)
	 */
	public CheckResultMessage check31(int day) {
		int r1 = get(39, 5);
		int c1 = get(40, 5);
		int r2 = get(43, 5);
		int c2 = get(44, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r2 + 3, c2 + day - 1, 11).add(
						getValue(r1 + 8, c1 + day - 1, 10).subtract(
								getValue(r1 + 15, c1 + day - 1, 10)).subtract(
								getValue(r1 + 18, c1 + day - 1, 10)));
				if (0 != getValue(r2 + 3, c2 + day, 11).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName
							+ ">备付金银行账户中未结转的支付业务净收入余额 M4:" + day + "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r2 + 3, c2 + day - 1, 11).add(
						getValue1(r1 + 8, c1 + day - 1, 10).subtract(
								getValue1(r1 + 15, c1 + day - 1, 10)).subtract(
								getValue1(r1 + 18, c1 + day - 1, 10)));
				if (0 != getValue1(r2 + 3, c2 + day, 11).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName
							+ ">备付金银行账户中未结转的支付业务净收入余额 M4:" + day + "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">备付金银行账户中未结转的支付业务净收入余额 M4:" + day
				+ "日正确");
	}

	/*
	 * 期末以现金形式持有的客户备付金余额 M5(T) = M5(T-1)+ L4 - L5
	 */
	public CheckResultMessage check32(int day) {
		int r1 = get(39, 5);
		int c1 = get(40, 5);
		int r2 = get(43, 5);
		int c2 = get(44, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r2 + 4, c2 + day - 1, 11).add(
						getValue(r1 + 4, c1 + day - 1, 10)).subtract(
						getValue(r1 + 5, c1 + day - 1, 10));
				if (0 != getValue(r2 + 4, c2 + day, 11).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName
							+ ">期末以现金形式持有的客户备付金余额 M5:" + day + "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r2 + 4, c2 + day - 1, 11).add(
						getValue1(r1 + 4, c1 + day - 1, 10)).subtract(
						getValue1(r1 + 5, c1 + day - 1, 10));
				if (0 != getValue1(r2 + 4, c2 + day, 11).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName
							+ ">期末以现金形式持有的客户备付金余额 M5:" + day + "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">期末以现金形式持有的客户备付金余额 M5:" + day
				+ "日正确");
	}

	/*
	 * 本期期末仍存在的以自有资金先行偿付的预付卡赎回金额M6(T) = M6(T-1) + (L6 - L7)
	 */
	public CheckResultMessage check33(int day) {
		int r1 = get(39, 5);
		int c1 = get(40, 5);
		int r2 = get(43, 5);
		int c2 = get(44, 5);
		BigDecimal b = new BigDecimal(0);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				b = getValue(r2 + 5, c2 + day - 1, 11).add(
						getValue(r1 + 6, c1 + day - 1, 10)).subtract(
						getValue(r1 + 7, c1 + day - 1, 10));
				if (0 != getValue(r2 + 5, c2 + day, 11).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName
							+ ">本期期末仍存在的以自有资金先行偿付的预付卡赎回金额M6:" + day + "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				b = getValue1(r2 + 5, c2 + day - 1, 11).add(
						getValue1(r1 + 6, c1 + day - 1, 10)).subtract(
						getValue1(r1 + 7, c1 + day - 1, 10));
				if (0 != getValue1(r2 + 5, c2 + day, 11).compareTo(b)) {
					return error("支付机构汇总报表<" + fileName
							+ ">本期期末仍存在的以自有资金先行偿付的预付卡赎回金额M6:" + day + "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">本期期末仍存在的以自有资金先行偿付的预付卡赎回金额M6:"
				+ day + "日正确");
	}

	/*
	 * 备付金账户中押金余额延续性 N1(T) = N6(T-1)
	 */
	public CheckResultMessage check34(int day) {
		int r = get(47, 5);
		int c = get(48, 5);
		if (checkVersion(file).equals("2003")) {
			try {
				in = new FileInputStream(file);
				hWorkbook = new HSSFWorkbook(in);
				if (0 != getValue(r + day, c, 12).compareTo(
						getValue(r + day - 1, c + 5, 12))) {
					return error("支付机构汇总报表<" + fileName + ">备付金账户中押金余额延续性 N1:"
							+ day + "日错误");
				}
			} catch (Exception e) {
			}
		} else {
			try {
				in = new FileInputStream(file);
				xWorkbook = new XSSFWorkbook(in);
				if (0 != getValue1(r + day, c, 12).compareTo(
						getValue1(r + day - 1, c + 5, 12))) {
					return error("支付机构汇总报表<" + fileName + ">备付金账户中押金余额延续性 N1:"
							+ day + "日错误");
				}
			} catch (Exception e) {
			}
		}
		return pass("支付机构汇总报表<" + fileName + ">备付金账户中押金余额延续性 N1:" + day + "日正确");
	}

	/**
	 * 返回正确信息
	 * 
	 * @param message
	 * @return
	 */
	public CheckResultMessage pass(String message) {
		return new CheckResultMessage(message);
	}

	/**
	 * 返回错误信息
	 * 
	 * @param message
	 * @return
	 */
	public CheckResultMessage error(String message) {
		return new CheckResultMessage(message, CheckResultMessage.CHECK_ERROR);
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
