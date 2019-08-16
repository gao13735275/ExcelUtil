package com.gp.office.excel.util.handler.write;

import com.gp.office.excel.util.model.write.CellRangeAddressDto;
import com.gp.office.excel.util.model.write.PictureDto;
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
     * 添加图片
     * @return
     */
    List<PictureDto> setPicture();
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
     * 设置单元格数据格式,默认是string，不传也是string,暂时只支持string和number,公式
     *
     * @throws IllegalArgumentException if the specified cell type is invalid
     * @see org.apache.poi.ss.usermodel.Cell#CELL_TYPE_NUMERIC
     * @see org.apache.poi.ss.usermodel.Cell#CELL_TYPE_STRING
     *
     * @param sheetIndex
     * @param rowIndex
     * @param columnIndex
     * @return
     */
    Integer setCellType(int sheetIndex, int rowIndex, int columnIndex);

    /**
     * 设置样式,如果先设置样式，然后单元格合并，则设置线条粗细，使用addMergedRegions方法里的CellRangeAddressDto设置线条
     * @return
     */
    CellStyle setCellStyle(int sheetIndex, int rowIndex, int columnIndex, Workbook wb);

    /**
     * 获取Workbook对象
     * @return
     */
    void getWorkbook(Workbook wb);
}
