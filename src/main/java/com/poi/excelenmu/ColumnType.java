package com.poi.excelenmu;

/**
 * 列类型
 * 
 * @author mbb
 *
 */
public enum ColumnType {

	simpleDateType(0, "yyyy-MM-dd HH:mm:ss"), ChineseDateType(1, "yyyy年MM月dd日 HH时mm分ss秒"), notDateType(2, "");
	private int key;
	private String value;

	private ColumnType(int key, String value) {
		this.key = key;
		this.value = value;
	}

	public int getKey() {
		return key;
	}

	public String getValue() {
		return value;
	}

	public void setValue(String value) {
		this.value = value;
	}

}
