package com.poi.anno;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(value = { ElementType.TYPE })
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface TableEntity {
	String tableName() default "tableName.xlsx";

	String sheetName() default "sheet1";

	int sheetIndex() default 0;

	int startIndex() default 2;

	String tableHead() default "";

}
