package com.gp.office.excel.util.handler.write;

import com.gp.office.excel.util.model.write.CellRangeAddressDto;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * @program: excelutil
 * @description:
 * @author: gaopo
 * @create: 2018-10-08 13:59
 **/
public interface OnWriterDataHandler {
    /**
     * 设置列宽
     * @return
     */
    List<Integer> setColumnWidth();

    /**
     * 设置行高
     * @return
     */
    Integer setRowHeight(int sheetIndex, int rowIndex);

    /**
     * 每行的数据
     * @param sheetIndex
     * @param rowIndex
     * @return
     */
    List<String> rowDataHandler(int sheetIndex, int rowIndex);

    /**
     * 总共几行
     * @param sheetIndex
     * @return
     */
    int rowCount(int sheetIndex);

    /**
     * 设置合并单元格的数据
     * @return
     */
    List<CellRangeAddressDto> addMergedRegions();

    /**
     * 标题
     * @return
     */
    String setSheetName();

    /**
     * 设置样式
     * @return
     */
    CellStyle setCellStyle(int sheetIndex, int rowIndex, int columnIndex, Workbook wb);

    /**
     * 获取Workbook对象
     * @return
     */
    void getWorkbook(Workbook wb);
}
