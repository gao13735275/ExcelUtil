package com.gp.office.excel.util.handler.read;

import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * @program: excelutil
 * @description:
 * @author: gaopo
 * @create: 2018-10-08 10:11
 **/
public interface OnReadDataHandler {
    /**
     * 每行的数据
     * @param sheetIndex
     * @param rowIndex
     * @param rowData
     */
    void rowDataHandler(int sheetIndex, int rowIndex, List<String> rowData);

    /**
     * 获取Workbook对象
     * @return
     */
    void getWorkbook(Workbook wb);

    /**
     * 总行数
     * @param sheetIndex
     * @param _rowCount
     */
    void rowCount(int sheetIndex, int _rowCount);
}
