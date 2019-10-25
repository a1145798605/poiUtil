package com.poi.entity;

import com.poi.anno.ExcelColumn;
import com.poi.anno.TableEntity;

@TableEntity(sheetIndex = 0, startIndex = 2)
public class Test01 {
	@ExcelColumn(columnIndex = 0, columnName = "id")
	private int id;
	@ExcelColumn(columnIndex = 1, columnName = "用户姓名")
	private String caseNo;
	@ExcelColumn(columnIndex = 2, columnName = "日期")
	private String shootTime;

	public int getId() {
		return id;
	}

	public void setId(int id) {
		this.id = id;
	}

	public String getCaseNo() {
		return caseNo;
	}

	public void setCaseNo(String caseNo) {
		this.caseNo = caseNo;
	}

	public String getShootTime() {
		return shootTime;
	}

	public void setShootTime(String shootTime) {
		this.shootTime = shootTime;
	}

	@Override
	public String toString() {
		return "Test01 [id=" + id + ", caseNo=" + caseNo + ", shootTime=" + shootTime + "]";
	}

}
