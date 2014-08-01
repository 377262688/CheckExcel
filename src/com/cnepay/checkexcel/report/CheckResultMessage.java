package com.cnepay.checkexcel.report;

public class CheckResultMessage {

	private String message;
	private int type;
	
	public final static int CHECK_OK = 0;
	public final static int CHECK_ERROR = 1;
	public final static int CHECK_WARN = 2;
	
	public final static int SYSTEM = 9;
	
	public CheckResultMessage() {
		this.message = "";
		this.type = CHECK_OK;
	}

	public CheckResultMessage(String message) {
		this.message = message;
		this.type = CHECK_OK;
	}
	
	public CheckResultMessage(String message, int type) {
		this.message = message;
		this.type = type;
	}	
	
	public String getMessage() {
		return message;
	}
	
	public void setMessage(String message) {
		this.message = message;
	}
	
	public int getType() {
		return type;
	}
	
	public void setType(int type) {
		this.type = type;
	}	
}
