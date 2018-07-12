package com.woter.fact.excel.exception;

/**
 * @author woter
 * @date 2017-9-29 下午7:04:38
 */
public class ExcelException extends RuntimeException {

    private static final long serialVersionUID = -5691872852547856320L;

    public ExcelException() {
        super();
    }

    public ExcelException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelException(String message) {
        super(message);
    }

    public ExcelException(Throwable cause) {
        super(cause);
    }

}
 