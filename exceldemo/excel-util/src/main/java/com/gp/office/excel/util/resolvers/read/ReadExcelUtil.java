package com.gp.office.excel.util.resolvers.read;

import com.gp.office.excel.util.constants.ExcelConstant;
import com.gp.office.excel.util.exception.ExcelException;
import com.gp.office.excel.util.handler.read.OnReadDataHandler;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @program: excelutil
 * @description:
 * @author: gaopo
 * @create: 2018-10-08 13:15
 **/
public class ReadExcelUtil {
    private final static Logger logger = LoggerFactory.getLogger(ReadExcelUtil.class);

    /**
     * 读取Excel数据(默认读取第一个工作表的所有数据,排除表头)
     * @param inputStream
     */
    public static void readExcel(InputStream inputStream, OnReadDataHandler handler) {
        readExcel(inputStream, 0, handler);
    }

    /**
     * 读取指定sheetIndex的数据（排除表头）
     * @param  inputStream
     * @param sheetIndex 工作表索引
     * @param handler 行数据处理回调
     */
    public static void readExcel(InputStream inputStream, int sheetIndex, OnReadDataHandler handler) {
        readExcel(inputStream, handler, sheetIndex, 1, -1, 0, -1);
    }

    /**
     * 读取Excel数据
     * @param inputStream
     * @param handler 行数据处理回调
     * @param sheetIndex 工作表索引
     * @param startRowIndex 开始行索引
     * @param endRowIndex 结束行索引,-1为所有
     * @param startCellIndex 开始列索引
     * @param endCellIndex 结束列索引,-1为所有
     */
    public static void readExcel(InputStream inputStream, OnReadDataHandler handler, int sheetIndex, int startRowIndex, int endRowIndex, int startCellIndex, int endCellIndex) {
        long totalRows = 0L;
        long begin = System.currentTimeMillis();
        Workbook wb = createWorkbook(inputStream);
        handler.getWorkbook(wb);//把workBook对象放出去

        if(wb != null) {
            Sheet sheet = wb.getSheetAt(sheetIndex);

            if(sheet != null) {

                if(endRowIndex == -1) {
                    endRowIndex = sheet.getPhysicalNumberOfRows();
                }
                if(endCellIndex == -1) {
                    endCellIndex = sheet.getRow(startRowIndex).getPhysicalNumberOfCells();
                }
                //总行数
                handler.rowCount(sheetIndex,endRowIndex);

                for (int i = startRowIndex; i < endRowIndex; i++) {
                    List<String> rowData = new ArrayList<String>();
                    Row row = sheet.getRow(i);
                    if(row != null) {
                        for (int j = startCellIndex; j < endCellIndex; j++) {
                            Cell cell = row.getCell(j);
                            if(cell != null) {
                                //时间型的数据
                                if (HSSFCell.CELL_TYPE_NUMERIC == cell.getCellType() && HSSFDateUtil.isCellDateFormatted(cell)) {
                                    rowData.add(cell.getDateCellValue().getTime()+"");
                                }else{
                                    // 统一以字符串的方式获取
                                    cell.setCellType(Cell.CELL_TYPE_STRING);
                                    if("".equals(cell.getStringCellValue())){
                                        rowData.add(ExcelConstant.getEmptyCellValue());
                                    }else{
                                        rowData.add(cell.getStringCellValue());
                                    }
                                }
                            } else {
                                rowData.add(ExcelConstant.getEmptyCellValue());
                            }
                        }
                    }
                    if(rowData.size() > 0) {
                        // 处理当前行数据
                        handler.rowDataHandler(sheetIndex,i,rowData);
                        totalRows++;
                    }
                }
            }
        }
        logger.info(String.format("Excel数据读取并处理完成,共读取数据：%s行,耗时：%s seconds.", totalRows,(System.currentTimeMillis() - begin) / 1000F));
    }

    protected static Workbook createWorkbook(InputStream inputStream) {
        if(!inputStream.markSupported()) {
            inputStream = new PushbackInputStream(inputStream, 8);//用退回流
        }
        try {
            if(POIFSFileSystem.hasPOIFSHeader(inputStream)){
                return new HSSFWorkbook(inputStream);//2003
            }

            if(POIXMLDocument.hasOOXMLHeader(inputStream)) {
                return new XSSFWorkbook(inputStream);//2003以上
            }
        } catch (IOException e) {
            throw new ExcelException("不能读取有效的Excel数据！");
        }

        return null;
//        Workbook workbook = null;
//        try {
//            String fileName =mulfile.getOriginalFilename();
//            String fileType = fileName.substring(fileName.lastIndexOf(".") + 1);
//            if(fileType.toLowerCase().equals("xls")){
//                workbook = new HSSFWorkbook(mulfile.getInputStream());//2003
//            }else{
//                workbook = new XSSFWorkbook(mulfile.getInputStream());//2007以上
//            }
//        } catch (IOException e) {
//        }
//        return workbook;
    }
}
