package com.gp.office.excel.util.handler.write;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * @program: excelutil
 * @description:
 * @author: gaopo
 * @create: 2018-10-08 17:10
 **/
public abstract class BaseOnWriterDataHandlerInternal implements OnWriterDataHandler {
    @Override
    public List<String> rowDataHandler(int sheetIndex, int rowIndex) {
//        if(sheetIndex > 0){
//            return null;
//        }

        if(rowIndex < tableDataTopRowCount(sheetIndex) ){//列表数据之前的数据
            return tableDataTop(sheetIndex,rowIndex);
        }else if(rowIndex >= tableDataTopRowCount(sheetIndex) && rowIndex < tableDataTopRowCount(sheetIndex) + tableDataRowCount(sheetIndex)){//列表数据
            return tableData(sheetIndex,rowIndex);
        }else{
            return tableDataDown(sheetIndex,rowIndex);
        }
    }

    @Override
    public int rowCount(int sheetIndex) {
        return tableDataTopRowCount(sheetIndex) + tableDataRowCount(sheetIndex) + tableDataDownRowCount(sheetIndex);
    }

    /**
     * 获取wb对象
     * @param wb
     */
    @Override
    public void getWorkbook(Workbook wb){

    }

    @Override
    public CellStyle setCellStyle(int sheetIndex, int rowIndex, int columnIndex, Workbook wb) {
        return null;
    }

    @Override
    public List<Integer> setColumnWidth() {
        return null;
    }

    @Override
    public Integer setRowHeight(int sheetIndex, int rowIndex) {
        return null;
    }

    /**
     * 列表数据之前的部分
     * @return
     */
    protected abstract List<String> tableDataTop(int sheetIndex, int rowIndex);

    /**
     * 列表数据
     * @return
     */
    protected abstract List<String> tableData(int sheetIndex, int rowIndex);

    /**
     * 列表数据下面的部分
     * @return
     */
    protected abstract List<String> tableDataDown(int sheetIndex, int rowIndex);

    /**
     * 列表数据之前部分的抬头
     * @return
     */
    protected abstract int tableDataTopRowCount(int sheetIndex);

    /**
     * 列表数据的行数
     * @return
     */
    protected abstract int tableDataRowCount(int sheetIndex);

    /**
     * 列表数据之后部分的抬头
     * @return
     */
    protected abstract int tableDataDownRowCount(int sheetIndex);
}
