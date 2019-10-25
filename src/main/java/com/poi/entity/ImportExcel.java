package com.poi.entity;

import java.util.Date;

import com.poi.anno.ExcelColumn;
import com.poi.anno.TableEntity;
import com.poi.excelenmu.ColumnType;

@TableEntity(sheetIndex = 0, tableName = "tttt.xlsx", sheetName = "test", tableHead = "用户管理")
public class ImportExcel {
	@ExcelColumn(columnIndex = 0, columnName = "用户姓名")
	private String userName;
	@ExcelColumn(columnIndex = 1, columnName = "用户年龄")
	private int age;
	@ExcelColumn(columnIndex = 2, columnName = "日期", columnType = ColumnType.ChineseDateType)
	private Date time;

	public String getUserName() {
		return userName;
	}

	public void setUserName(String userName) {
		this.userName = userName;
	}

	public int getAge() {
		return age;
	}

	public void setAge(int age) {
		this.age = age;
	}

	public Date getTime() {
		return time;
	}

	public void setTime(Date time) {
		this.time = time;
	}
}
