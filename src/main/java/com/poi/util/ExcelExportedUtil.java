package com.poi.util;

import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.poi.anno.ExcelColumn;
import com.poi.anno.TableEntity;
import com.poi.entity.MyExcelException;
import com.poi.excelenmu.ColumnType;
import com.poi.excelenmu.ColumnValiDataIsNull;
import com.poi.excelenmu.ExcelType;
import com.poi.staticconst.StaticConst;

public class ExcelExportedUtil<E> {

	private ExcelExportedUtil() {
		throw new RuntimeException("无法进行实例化操作");
	}

	/**
	 * 导出到 excel表格
	 * 
	 * @param list
	 * @throws Exception
	 */
	public static <E> void exportExcel(OutputStream out, List<E> list) throws Exception {
		if (out == null) {
			System.out.println("文件不能为空");
			return;
		}
		if (list == null) {
			System.out.println("参数不能为空");
			return;
		}
		if (list.size() == 0) {
			System.out.println("列表无数据");
			return;
		}
		E e = list.get(0);
		Class<?> class1 = e.getClass();
		TableEntity tableEntity = class1.getAnnotation(TableEntity.class);
		if (tableEntity != null) {
			// 得到从第几行插入
			Map<String, Object> tableMap = Maps.newHashMap();
			int startIndex = tableEntity.startIndex();
			String sheetName = tableEntity.sheetName();
			String tableHead = tableEntity.tableHead();
			String tableName = tableEntity.tableName();
			tableMap.put("startIndex", startIndex);
			tableMap.put("sheetName", sheetName);
			tableMap.put("tableHead", tableHead);
			tableMap.put("tableName", tableName);
			// 解析column
			Field[] declaredFields = class1.getDeclaredFields();
			/**
			 * 列 和 field 的对应
			 */
			Map<Integer, String> map = Maps.newHashMap();
			/**
			 * 列和列的头
			 */
			Map<Integer, String> tableHeadMap = Maps.newHashMap();
			/**
			 * 列和 类型的对应
			 */
			Map<Integer, String> entityType = Maps.newHashMap();
			/**
			 * 处理日期类型
			 */
			Map<Integer, ColumnType> dateMap = Maps.newHashMap();
			/**
			 * 处理数据 格式校验
			 */
			Map<Integer, ColumnValiDataIsNull> dataValiedMap = Maps.newHashMap();
			getFiledMap(declaredFields, map, tableHeadMap, entityType, dateMap, dataValiedMap);
			startExportExcel(out, class1, list, tableMap, map, tableHeadMap, entityType, dateMap, dataValiedMap);
		}
	}

	public static ExcelType getExcelType(String fileName) {
		if (fileName == null || fileName.length() == 0) {
			return ExcelType.ExcelError;
		}
		if (fileName.endsWith(StaticConst.EXCEL_XLS)) {
			return ExcelType.ExcelType03;
		}
		if (fileName.endsWith(StaticConst.EXCEL_XLSX)) {
			return ExcelType.ExcelType07;
		}
		return ExcelType.ExcelError;
	}

	private static <E> void startExportExcel(OutputStream out, Class<?> clzz, List<E> list,
			Map<String, Object> tableMap, Map<Integer, String> map, Map<Integer, String> tableHeadMap,
			Map<Integer, String> entityType, Map<Integer, ColumnType> dateMap,
			Map<Integer, ColumnValiDataIsNull> dataValiedMap) throws Exception {
		Workbook workbook = getWorkBook(null, getExcelType(tableMap.get("tableName").toString())); // 创建工作簿对象
		Sheet sheet = workbook.createSheet(tableMap.get("sheetName").toString()); // 创建工作表
		// 判断是否需要产生标题
		String tableHead = tableMap.get("tableHead").toString();
		if (tableHead != null && tableHead.length() > 0) {
			// 需要产生头
			Row rowm = sheet.createRow(0); // 产生表格标题行
			Cell cellTiltle = rowm.createCell(0);
			CellStyle columnTopStyle = getColumnTopStyle(workbook);
			sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, map.size() - 1));
			cellTiltle.setCellStyle(columnTopStyle); // 设置标题行样式
			cellTiltle.setCellValue(tableHead); // 设置标题行值
			// 没有 标题头的情况
			setExcelHeadTitle(list, workbook, clzz, sheet, map, tableHeadMap, entityType, dateMap, dataValiedMap, 1);
		} else {
			setExcelHeadTitle(list, workbook, clzz, sheet, map, tableHeadMap, entityType, dateMap, dataValiedMap, 0);
		}
		workbook.write(out);
	}

	private static <E> void setExcelHeadTitle(List<E> list, Workbook workbook, Class<?> clzz, Sheet sheet,
			Map<Integer, String> map, Map<Integer, String> tableHeadMap, Map<Integer, String> entityType,
			Map<Integer, ColumnType> dateMap, Map<Integer, ColumnValiDataIsNull> dataValiedMap, int index)
			throws Exception {
		// 没有 标题头的情况
		if (tableHeadMap != null && tableHeadMap.size() > 0) {
			Set<Integer> keySet = tableHeadMap.keySet();
			// 创建 第二行 设置标题头
			// 将列头设置到sheet的单元格中
			Row rowRowName = sheet.createRow(index); // 在索引1的位置创建行(最顶端的行开始的第二行)
			for (Integer key : keySet) {
				Cell createCell = rowRowName.createCell(key);
				createCell.setCellStyle(getStyle(workbook));
				createCell.setCellValue(tableHeadMap.get(key));
			}
			setExcelTableValue(list, workbook, clzz, sheet, map, tableHeadMap, entityType, dateMap, dataValiedMap,
					index + 1);
		} else {
			setExcelTableValue(list, workbook, clzz, sheet, map, tableHeadMap, entityType, dateMap, dataValiedMap,
					index);
		}
	}

	private static <E> void setExcelTableValue(List<E> list, Workbook workbook, Class<?> clzz, Sheet sheet,
			Map<Integer, String> map, Map<Integer, String> tableHeadMap, Map<Integer, String> entityType,
			Map<Integer, ColumnType> dateMap, Map<Integer, ColumnValiDataIsNull> dataValiedMap, int index)
			throws Exception {
		for (int i = 0; i < list.size(); i++) {
			// 取出当前的对象
			Object object = list.get(i);
			Row row = sheet.createRow(i + index);
			Set<Integer> keySet2 = map.keySet();
			for (Integer key : keySet2) {
				Cell filed = row.createCell(key);
				filed.setCellStyle(getStyle(workbook));
				// 列名
				String fieldName = map.get(key);
				Field declaredField = clzz.getDeclaredField(fieldName);
				declaredField.setAccessible(true);
				Object cellvalue = declaredField.get(object);
				String columnType = entityType.get(key);
				setCellValue(filed, columnType, cellvalue, key, dateMap, dataValiedMap);
			}
		}
	}

	private static void setCellValue(Cell cell, String columnType, Object cellvalue, Integer key,
			Map<Integer, ColumnType> dateMap, Map<Integer, ColumnValiDataIsNull> dataValiedMap) throws Exception {
		ColumnValiDataIsNull columnValiDataIsNull = dataValiedMap.get(key);
		if (columnValiDataIsNull.getKey() == 1) {
			// 需要验证是否需要 做 null值校验
			if (cellvalue == null) {
				throw new MyExcelException("excel数据输入异常，数据不能为空 数据类型为：" + dateMap.get(key) + "第" + key + "个对象");
			}
		}
		switch (columnType) {
		case "Double":
		case "double":

			cell.setCellValue(Double.parseDouble(cellvalue.toString()));
			break;
		case "String":
			cell.setCellType(CellType.STRING);
			cell.setCellValue(cellvalue.toString());
			break;
		case "int":
		case "Integer":
			cell.setCellType(CellType.STRING);
			cell.setCellValue(cellvalue.toString());
			break;
		case "Boolean":
		case "boolean":
			cell.setCellValue(Boolean.parseBoolean(cellvalue.toString()));
			break;
		case "Date":
			ColumnType dateType = dateMap.get(key);
			SimpleDateFormat sim = new SimpleDateFormat(dateType.getValue());
			String format = sim.format(cellvalue);
			cell.setCellValue(format);
			break;
		}
	}

	private static CellStyle getColumnTopStyle(Workbook workbook) {
		Font font = workbook.createFont();
		// 设置字体大小
		font.setFontHeightInPoints((short) 11);
		// 字体加粗
		font.setBold(true);
		// 设置字体名字
		font.setFontName("Courier New");
		// 设置样式;
		CellStyle style = workbook.createCellStyle();
		// 设置底边框;
		style.setBorderBottom(BorderStyle.THIN);
		// 设置底边框颜色;
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		// 设置左边框;
		style.setBorderLeft(BorderStyle.THIN);
		// 设置左边框颜色;
		style.setLeftBorderColor(IndexedColors.BLACK.index);
		// 设置右边框;
		style.setBorderRight(BorderStyle.THIN);
		// 设置右边框颜色;
		style.setRightBorderColor(IndexedColors.BLACK.index);
		// 设置顶边框;
		style.setBorderTop(BorderStyle.THIN);
		// 设置顶边框颜色;
		style.setTopBorderColor(IndexedColors.BLACK.index);
		// 在样式用应用设置的字体;
		style.setFont(font);
		// 设置自动换行;
		style.setWrapText(false);
		// 设置水平对齐的样式为居中对齐;
		style.setAlignment(HorizontalAlignment.CENTER);
		// 设置垂直对齐的样式为居中对齐;
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		return style;

	}

	/*
	 * 列数据信息单元格样式
	 */
	private static CellStyle getStyle(Workbook workbook) {
		// 设置字体
		Font font = workbook.createFont();
		// 设置字体大小
		// font.setFontHeightInPoints((short)10);
		// 字体加粗
		// font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		// 设置字体名字
		font.setFontName("Courier New");
		// 设置样式;
		CellStyle style = workbook.createCellStyle();
		// 设置底边框;
		style.setBorderBottom(BorderStyle.THIN);
		// 设置底边框颜色;
		style.setBottomBorderColor(IndexedColors.BLUE.index);
		// 设置左边框;
		style.setBorderLeft(BorderStyle.THIN);
		// 设置左边框颜色;
		style.setLeftBorderColor(IndexedColors.BLACK.index);
		// 设置右边框;
		style.setBorderRight(BorderStyle.THIN);
		// 设置右边框颜色;
		style.setRightBorderColor(IndexedColors.BLACK.index);
		// 设置顶边框;
		style.setBorderTop(BorderStyle.THIN);
		// 设置顶边框颜色;
		style.setTopBorderColor(IndexedColors.BLACK.index);
		// 在样式用应用设置的字体;
		style.setFont(font);
		// 设置自动换行;
		style.setWrapText(false);
		// 设置水平对齐的样式为居中对齐;
		style.setAlignment(HorizontalAlignment.CENTER);
		// 设置垂直对齐的样式为居中对齐;
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		return style;
	}

	private static void getFiledMap(Field[] declaredFields, Map<Integer, String> map, Map<Integer, String> tableHeadMap,
			Map<Integer, String> entityType, Map<Integer, ColumnType> dateMap,
			Map<Integer, ColumnValiDataIsNull> dataValiedMap) {
		for (Field field : declaredFields) {
			if (field != null) {
				field.setAccessible(true);
				ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
				Class<?> type = field.getType();
				String columnName = annotation.columnName();
				int columnIndex = annotation.columnIndex();
				String name = field.getName();
				entityType.put(columnIndex, type.getSimpleName());
				if ("Date".equals(type.getSimpleName())) {
					ColumnType columnType = annotation.columnType();
					dateMap.put(columnIndex, columnType);
				}
				ColumnValiDataIsNull valiData = annotation.valiData();
				// 为0不需要 判空
				dataValiedMap.put(columnIndex, valiData);
				if (tableHeadMap != null) {
					tableHeadMap.put(columnIndex, columnName);
				}
				map.put(columnIndex, name);
			}
		}
	}

	private static void getFiledMap(Field[] declaredFields, Map<Integer, String> map, Map<Integer, String> entityType,
			Map<Integer, ColumnType> dateMap, Map<Integer, ColumnValiDataIsNull> dataValiedMap) {
		getFiledMap(declaredFields, map, null, entityType, dateMap, dataValiedMap);
	}

	// 不判断继承的情况 不考虑继承 以后在考虑继承
	public static <E> List<E> importExcel(InputStream in, Class<?> clzz, ExcelType excelType) throws Exception {
		// 1 第一步 判断是否是 tableentity实体
		TableEntity tableEntity = clzz.getAnnotation(TableEntity.class);
		if (tableEntity != null) {
			// 得到要开始插入的sheetIndex
			int sheetIndex = tableEntity.sheetIndex();
			// 得到从第几行插入
			int startIndex = tableEntity.startIndex();
			// 解析column
			Field[] declaredFields = clzz.getDeclaredFields();
			/**
			 * 列 和 field 的对应
			 */
			Map<Integer, String> fieldmap = Maps.newHashMap();
			/**
			 * 列和 类型的对应
			 */
			Map<Integer, String> entityType = Maps.newHashMap();
			/**
			 * 日期对应格式
			 */
			Map<Integer, ColumnType> dateMap = Maps.newHashMap();
			/**
			 * 数据校验
			 */
			Map<Integer, ColumnValiDataIsNull> dataValiedMap = Maps.newHashMap();
			getFiledMap(declaredFields, fieldmap, entityType, dateMap, dataValiedMap);
			return startImportExcel(in, clzz, sheetIndex, startIndex, fieldmap, excelType, entityType, dateMap,
					dataValiedMap);
		} else {
			return null;
		}

	}

	private static <E> List<E> startImportExcel(InputStream in, Class<?> clzz, int sheetIndex, int startIndex,
			Map<Integer, String> map, ExcelType excelType, Map<Integer, String> entityType,
			Map<Integer, ColumnType> dateMap, Map<Integer, ColumnValiDataIsNull> dataValiedMap) throws Exception {
		List<E> result = Lists.newArrayList();
		Workbook workBook = getWorkBook(in, excelType);
		Sheet sheet = workBook.getSheetAt(sheetIndex);
		if (sheet == null) {
			return result;
		}
		for (int rowIndex = startIndex; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
			// 获得当前行
			Row row = sheet.getRow(rowIndex);
			if (row == null) {
				continue;
			}
			// new 出当前实体类
			Object newInstance = clzz.newInstance();
			for (int index = 0; index < row.getLastCellNum(); index++) {
				Cell cell = row.getCell(index);
				String indexName = map.get(index);
				String columnType = entityType.get(index);
				Field declaredField = clzz.getDeclaredField(indexName);
				declaredField.setAccessible(true);
				// 获取当前行的数据 要判断类型
				Object cellValue = getCellValue(cell, columnType, dataValiedMap, index, rowIndex);
				declaredField.set(newInstance, cellValue);
			}
			result.add((E) newInstance);
			// 获得当前行的每一列的列名

		}
		return result;
	}

	private static Object getCellValue(Cell cell, String cellType, Map<Integer, ColumnValiDataIsNull> dataValiedMap,
			int index, int rowIndex) {
		try {
			boolean flag = false;
			if (dataValiedMap.get(index).getKey() == 1) {
				// 需要验证
				flag = true;
			}
			switch (cellType) {
			case "Double":
			case "double":
				if (flag) {
					if (cell == null) {
						// 进行容忍
						throw new MyExcelException(
								"数据不能为空！！！！第" + (rowIndex + 1) + "行" + "中的，第" + (index + 1) + "列，数据为空");
					}
				} else {
					if (cell == null) {
						// 进行容忍
						return 0.0;
					}
				}
				return cell.getNumericCellValue();
			case "String":

				if (flag) {
					if (cell == null) {
						throw new MyExcelException(
								"数据不能为空！！！！第" + (rowIndex + 1) + "行" + "中的，第" + (index + 1) + "列，数据为空");
					}
					if (cell.getStringCellValue() != null && cell.getStringCellValue().trim().length() > 0) {
						return cell.getStringCellValue();
					} else {
						throw new MyExcelException(
								"数据不能为空！！！！第" + (rowIndex + 1) + "行" + "中的，第" + (index + 1) + "列，数据为空");
					}
				} else {
					if (cell == null) {
						return "";
					}
					return cell.getStringCellValue();
				}

			case "Date":
				if (flag) {
					if (cell == null) {
						throw new MyExcelException(
								"数据不能为空！！！！第" + (rowIndex + 1) + "行" + "中的，第" + (index + 1) + "列，数据为空");
					}
					if (cell.getDateCellValue() != null) {
						return cell.getDateCellValue();
					}
					throw new MyExcelException("数据不能为空！！！！第" + (rowIndex + 1) + "行" + "中的，第" + (index + 1) + "列，数据为空");

				} else {
					if (cell == null) {
						return new Date();
					} else {
						return cell.getDateCellValue();
					}
				}

			case "Boolean":
			case "boolean":
				if (flag) {
					if (cell == null) {
						throw new MyExcelException(
								"数据不能为空！！！！第" + (rowIndex + 1) + "行" + "中的，第" + (index + 1) + "列，数据为空");
					}
				} else {
					if (cell == null) {
						return true;
					}
				}
				return cell.getBooleanCellValue();
			case "int":
			case "Integer":
				if (flag) {
					if (cell == null) {
						throw new MyExcelException(
								"数据不能为空！！！！第" + (rowIndex + 1) + "行" + "中的，第" + (index + 1) + "列，数据为空");
					}
				} else {
					if (cell == null) {
						return 0;
					}
				}
				String split = (cell.getNumericCellValue() + "").split("\\.")[0];
				return Integer.parseInt(split);
			}

		} catch (IllegalStateException e) {
			throw new MyExcelException(
					"类型转换异常" + e.getMessage() + "第" + (rowIndex + 1) + "行" + "中的，第" + (index + 1) + "列");
		}
		return "";
	}

	private static Workbook getWorkBook(InputStream is, ExcelType excelType) throws Exception {
		Workbook workbook = null;
		if (is != null) {
			if (excelType.getKey() == 0) {
				workbook = new HSSFWorkbook(is);
			} else if (excelType.getKey() == 1) {
				workbook = new XSSFWorkbook(is);
			}
		} else {
			if (excelType.getKey() == 0) {
				workbook = new HSSFWorkbook();
			} else if (excelType.getKey() == 1) {
				workbook = new XSSFWorkbook();
			}
		}
		return workbook;
	}

}
