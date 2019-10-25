package com.poi.excelenmu;

public enum ColumnValiDataIsNull {
	CanNotEmpty(1, "不能为空"), CanEmpty(0, "可以为空");
	private int key;
	private String value;

	private ColumnValiDataIsNull(int key, String value) {
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
