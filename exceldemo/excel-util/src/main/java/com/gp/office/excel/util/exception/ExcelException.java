package com.gp.office.excel.util.exception;

/**
 * @program: excelutil
 * @description:
 * @author: gaopo
 * @create: 2018-10-08 10:08
 **/
public class ExcelException extends RuntimeException {

    private static final long serialVersionUID = -5621100658541541518L;
    protected String logError=null;
    protected Integer code = -1;

    public String getLogError() {
        return logError;
    }

    public Integer getCode(){
        return this.code;
    }

    private ExcelException() {}

    public ExcelException(String message) {
        super(message);
    }

    public ExcelException(String message, String _logError){
        this(message);
        logError = _logError;
    }

    public ExcelException(Integer code, String message) {
        super(message);
        this.code = code;
    }

    public ExcelException(Integer code, String message, String _logError){
        this(code,message);
        logError = _logError;
    }
}
