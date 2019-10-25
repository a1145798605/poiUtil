package com.poi.Main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import com.google.common.collect.Lists;
import com.poi.entity.ExportedEntity;
import com.poi.entity.ImportExcel;
import com.poi.excelenmu.ExcelType;
import com.poi.util.ExcelExportedUtil;

public class TestMain {
	List<ExportedEntity> test = new ArrayList<ExportedEntity>();

	public static void main(String[] args) throws Exception {

		File file = new File("d:/test.xlsx");
		FileInputStream in = new FileInputStream(file);
		List<ExportedEntity> importExcel = ExcelExportedUtil.importExcel(in, ExportedEntity.class,
				ExcelType.ExcelType07);
		System.out.println(importExcel);
		File test = new File("d:/kk/test.xlsx");
		ImportExcel e1 = new ImportExcel();
		e1.setAge(100);
		e1.setTime(new Date());
		e1.setUserName("test");
		List<ImportExcel> list = Lists.newArrayList(e1);
		OutputStream out = new FileOutputStream(test);
		ExcelExportedUtil.exportExcel(out, list);
	}

}
