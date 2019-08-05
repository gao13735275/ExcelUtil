package com.gp.office.excel.util;

import com.gp.office.excel.util.constants.ExcelType;
import com.gp.office.excel.util.handler.write.OnWriterDataHandlerInternalImpl;
import com.gp.office.excel.util.resolvers.write.WriteExcelUtil;

import java.io.OutputStream;
import java.util.List;

/**
 * @program: ourHouse
 * @description:
 * @author: gaopo
 * @create: 2019-06-20 15:28
 **/
public class WriteExcelApi {
    /**
     * 注解形式导出excel
     * @param stream
     * @param sheetName
     * @param list
     * @param classOfT
     * @param <T>
     */
    public static <T> void writerExcelByAnnotation(OutputStream stream,String sheetName, List<?> list,Class<T> classOfT) {
        WriteExcelUtil writeExcelUtil = new WriteExcelUtil(ExcelType.EXCEL2007ByBulkData);
        writeExcelUtil.writerExcel(new OnWriterDataHandlerInternalImpl(sheetName,list,classOfT));
        writeExcelUtil.write(stream);
    }

    /**
     * 注解形式导出excel
     * @param stream
     * @param sheetName
     * @param _banner 大的标题
     * @param list
     * @param classOfT
     * @param <T>
     */
    public static <T> void writerExcelByAnnotation(OutputStream stream,String sheetName,String _banner, List<?> list,Class<T> classOfT) {
        WriteExcelUtil writeExcelUtil = new WriteExcelUtil(ExcelType.EXCEL2007ByBulkData);
        writeExcelUtil.writerExcel(new OnWriterDataHandlerInternalImpl(sheetName,_banner,list,classOfT));
        writeExcelUtil.write(stream);
    }

}
