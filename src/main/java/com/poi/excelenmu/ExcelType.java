package com.poi.excelenmu;

public enum ExcelType {

	ExcelType03(0, "03版本的excel格式以xls结尾"), 
	ExcelType07(1, "07版本的excel格式以xls结尾"), 
	ExcelError(2, "不正确的文件名字");
	private int key;
	private String value;

	private ExcelType(int key, String value) {
		this.key = key;
		this.value = value;
	}

	public int getKey() {
		return key;
	}

	public void setKey(int key) {
		this.key = key;
	}

	public String getValue() {
		return value;
	}

	public void setValue(String value) {
		this.value = value;
	}

}
