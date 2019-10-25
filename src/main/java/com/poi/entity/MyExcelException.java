package com.poi.entity;

public class MyExcelException extends RuntimeException {

	public MyExcelException() {
		super();
	}

	public MyExcelException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
		super(message, cause, enableSuppression, writableStackTrace);
	}

	public MyExcelException(String message, Throwable cause) {
		super(message, cause);
	}

	public MyExcelException(String message) {
		super(message);
	}

	public MyExcelException(Throwable cause) {
		super(cause);
	}

}
