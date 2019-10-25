package com.poi.entity;

import java.util.Date;

import com.poi.anno.ExcelColumn;
import com.poi.anno.TableEntity;
import com.poi.excelenmu.ColumnType;
import com.poi.excelenmu.ColumnValiDataIsNull;

@TableEntity(startIndex = 2, sheetIndex = 0)
public class ExportedEntity {

	@ExcelColumn(columnIndex = 0, columnName = "用户姓名")
	private String userName;
	@ExcelColumn(columnIndex = 1, columnName = "用户姓名")
	private int age;
	@ExcelColumn(columnIndex = 2, columnName = "日期", columnType = ColumnType.simpleDateType, valiData = ColumnValiDataIsNull.CanNotEmpty)
	private Date time;

	public Date getTime() {
		return time;
	}

	public void setTime(Date time) {
		this.time = time;
	}

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

	@Override
	public String toString() {
		return "ImportEntity [userName=" + userName + ", age=" + age + ", time=" + time + "]";
	}

}
