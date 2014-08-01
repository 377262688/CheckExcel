package com.cnepay.checkexcel.report;

import jxl.biff.CellReferenceHelper;

public class DayCell {
	
	private int startCol;
	private int startRow;
	private int direction;
	
	public final static int COL_DIRECTION = 0;
	public final static int ROW_DIRECTION = 1;
	
	public DayCell(int startCol, int startRow, int direction) {
		this.startCol = startCol;
		this.startRow = startRow;
		this.direction = direction;
	}
	
	public DayCell(String cellRef, int direction) {
		this.startCol = CellReferenceHelper.getColumn(cellRef);
		this.startRow = CellReferenceHelper.getRow(cellRef);
		this.direction = direction;
	}
	
	public int getStartCol() {
		return startCol;
	}
	public void setStartCol(int startCol) {
		this.startCol = startCol;
	}
	public int getStartRow() {
		return startRow;
	}
	public void setStartRow(int startRow) {
		this.startRow = startRow;
	}
	public int getDirection() {
		return direction;
	}
	public void setDirection(int direction) {
		this.direction = direction;
	}
	
	
}
