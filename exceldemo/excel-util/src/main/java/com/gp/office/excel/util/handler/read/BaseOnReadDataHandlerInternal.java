package com.gp.office.excel.util.handler.read;

import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * @program: excelutil
 * @description:
 * @author: gaopo
 * @create: 2018-10-09 11:12
 **/
public abstract class BaseOnReadDataHandlerInternal implements OnReadDataHandler {
    private int rowCount = -1;

    /**
     * 每行的数据
     * @param sheetIndex
     * @param rowIndex
     * @param rowData
     */
    @Override
    public void rowDataHandler(int sheetIndex, int rowIndex, List<String> rowData) {
        if(sheetIndex > 0){
            return;
        }

        if(rowIndex < tableDataTopRowCount(sheetIndex) ){//列表数据之前的数据
            tableDataTop(sheetIndex,rowIndex,rowData);
        }else if(rowIndex >= tableDataTopRowCount(sheetIndex) && rowIndex < rowCount - tableDataDownRowCount(sheetIndex)){//列表数据
            tableData(sheetIndex,rowIndex,rowData);
        }else{
            tableDataDown(sheetIndex,rowIndex,rowData);
        }
    }

    /**
     * 获取Workbook对象
     * @param wb
     */
    @Override
    public void getWorkbook(Workbook wb) {

    }

    /**
     * 总行数
     * @param sheetIndex
     * @param _rowCount
     */
    @Override
    public void rowCount(int sheetIndex,int _rowCount){
        rowCount = _rowCount;
    }

    /**
     * 列表数据之前的部分
     * @param sheetIndex
     * @param rowIndex
     * @param rowData
     * @return
     */
    protected abstract void tableDataTop(int sheetIndex, int rowIndex,List<String> rowData);

    /**
     * 列表数据
     * @param sheetIndex
     * @param rowIndex
     * @param rowData
     */
    protected abstract void tableData(int sheetIndex, int rowIndex,List<String> rowData);

    /**
     * 列表数据下面的部分
     * @param sheetIndex
     * @param rowIndex
     * @param rowData
     */
    protected abstract void tableDataDown(int sheetIndex, int rowIndex,List<String> rowData);

    /**
     * 列表数据之前部分的行数
     * @return
     */
    protected abstract int tableDataTopRowCount(int sheetIndex);

    /**
     * 列表数据之后部分的行数
     * @return
     */
    protected abstract int tableDataDownRowCount(int sheetIndex);
}
