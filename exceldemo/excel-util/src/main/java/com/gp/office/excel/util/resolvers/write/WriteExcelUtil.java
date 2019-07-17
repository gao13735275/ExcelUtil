package com.gp.office.excel.util.resolvers.write;

import com.gp.office.excel.util.constants.ExcelConstant;
import com.gp.office.excel.util.constants.ExcelType;
import com.gp.office.excel.util.handler.write.OnWriterDataHandler;
import com.gp.office.excel.util.model.write.CellRangeAddressDto;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.OutputStream;
import java.util.List;

/**
 * @program: excelutil
 * @description:
 * @author: gaopo
 * @create: 2018-10-08 17:22
 **/
public class WriteExcelUtil {
    private static Logger log = LoggerFactory.getLogger(WriteExcelUtil.class);

    /**
     * 生成excel
     *
     * @param stream
     * @param handler
     */
    public static void writerExcel(OutputStream stream, OnWriterDataHandler handler) {
        writerExcel(stream, ExcelType.EXCEL2007, handler);
    }

    /**
     * 生成大批量数据的导出
     * @param stream
     * @param handler
     */
    public static void writerExcelByBulkData(OutputStream stream, OnWriterDataHandler handler){
        writerExcel(stream, ExcelType.EXCEL2007ByBulkData, handler);
    }

    /**
     * 生成excel
     *
     * @param stream
     * @param handler
     */
    public static void writerExcel(OutputStream stream,int excelType, OnWriterDataHandler handler) {
        long totalRows = 1L;
        long begin = System.currentTimeMillis();

        Workbook wb = createWorkbook(excelType);
        handler.getWorkbook(wb);//把workBook对象放出去

        String sheetName = handler.setSheetName();
        if (sheetName.isEmpty()) {//如果没有sheet，则认为数据为空
            return;
        }
        //初始化表格
        Sheet sheet = wb.createSheet(sheetName);
        int sheetIndex = wb.getSheetIndex(sheet);

        log.debug("set excel columnWidth");
        List<Integer> setColumnWidth = handler.setColumnWidth();//以一个字符的1/256的宽度作为一个单位,3000就是11.7,就是约等于11个字符,这里说的都是半角
        if (setColumnWidth != null) {
            for (int i = 0; i < setColumnWidth.size(); i++) {
                sheet.setColumnWidth(i, setColumnWidth.get(i));
            }
        }


        int rowCount = handler.rowCount(sheetIndex);
        //填充数据
        log.debug("set excel row");
        for (int i = 0; i < rowCount; i++) {
            Row row = sheet.createRow(i);
            totalRows++;
            log.debug("get row data current row is {}",i);
            List<String> list = handler.rowDataHandler(sheetIndex, i);//每一行的数据
            if(list == null || list.size() == 0){
                continue;
            }

            log.debug("set rowHeight current row is {}",i);
            Integer rowHeight = handler.setRowHeight(sheetIndex,i);//设置行高
            if(rowHeight!=null && !rowHeight.equals(0)){
                row.setHeight(rowHeight.shortValue());
            }

            log.debug("while cell current row is {}",i);
            //循环一行里的所有列
            for (int j = 0; j < list.size(); j++) {
                Cell cell = row.createCell(j);
                String value = list.get(j).trim();
                cell.setCellType(Cell.CELL_TYPE_STRING);//设置所有数据都是string

                log.debug("set cell style,current row is {},current cell is {}",i,j);
                CellStyle cellStyle1 = handler.setCellStyle(sheetIndex,i,j,wb);
                if (cellStyle1 != null) {
                    cell.setCellStyle(cellStyle1);
                }

                log.debug("set cell data,current row is {},current cell is {}",i,j);
                if (!ExcelConstant.getEmptyCellValue().equals(value)) {
                    cell.setCellValue(value);
                }
            }
        }

        log.debug("excel merge cell");
        //合并单元格
        List<CellRangeAddressDto> cellRangeAddressDtoList = handler.addMergedRegions();
        if(cellRangeAddressDtoList != null && cellRangeAddressDtoList.size() != 0){
            for (CellRangeAddressDto cellRangeAddressDto : cellRangeAddressDtoList) {
                sheet.addMergedRegion(cellRangeAddressDto);
            }
        }

        try {
            log.debug("write to excel");
            //写excel
            wb.write(stream);
            stream.flush();
            stream.close();
        } catch (Exception ex) {
            log.error("导出excel出错:{}",ex);
        }
        log.info(String.format("Excel数据写入并处理完成,共写入数据：%s行,耗时：%s seconds.", totalRows, (System.currentTimeMillis() - begin) / 1000F));
    }

    /**
     * 创建一个workBook
     *
     * @param type
     * @return
     */
    private static Workbook createWorkbook(int type) {
        if (type == ExcelType.EXCEL2003){
            return new HSSFWorkbook();
        }else if(type == ExcelType.EXCEL2007ByBulkData){
            return new SXSSFWorkbook();
        }
        return new XSSFWorkbook();
    }

}
