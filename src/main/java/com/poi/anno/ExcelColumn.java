package com.poi.anno;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import com.poi.excelenmu.ColumnType;
import com.poi.excelenmu.ColumnValiDataIsNull;

@Target(value = { ElementType.FIELD })
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelColumn {
	/**
	 * 导出的 列的名字
	 * 
	 * @return
	 */
	String columnName();

	/**
	 * 对应的 列的下标
	 * 
	 * @return
	 */
	int columnIndex();

	/**
	 * 日期的类型
	 * 
	 * @return
	 */
	ColumnType columnType() default ColumnType.notDateType;

	/**
	 * 是否验证数据 是不是可以为空
	 */
	ColumnValiDataIsNull valiData() default ColumnValiDataIsNull.CanEmpty;
}
