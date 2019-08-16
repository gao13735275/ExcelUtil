package com.gp.office.excel.util.resolvers.write;

import com.gp.office.excel.util.constants.ExcelConstant;
import com.gp.office.excel.util.constants.ExcelType;
import com.gp.office.excel.util.exception.ExcelException;
import com.gp.office.excel.util.handler.write.OnWriterDataHandler;
import com.gp.office.excel.util.model.write.CellRangeAddressDto;
import com.gp.office.excel.util.model.write.PictureDto;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.RegionUtil;
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
    private Logger log = LoggerFactory.getLogger(WriteExcelUtil.class);

    Workbook wb;
    int excelType;

    /**
     * 初始化
     * @param _excelType
     */
    public WriteExcelUtil(int _excelType){//ExcelType.EXCEL2007ByBulkData 大批量导出
        this.createWorkbook(_excelType);
        excelType = _excelType;
    }

    /**
     * 组装完数据后，写入流并导出
     * @param stream
     */
    public void write(OutputStream stream){
        try {
            log.debug("write to excel");
            //写excel
            wb.write(stream);
            stream.flush();
            stream.close();
        } catch (Exception ex) {
            log.error("导出excel出错:{}",ex);
        }
    }


    /**
     * 生产excel
     * @param handlerList
     * @return
     */
    public WriteExcelUtil writerExcel(List<OnWriterDataHandler> handlerList){
        for (OnWriterDataHandler onWriterDataHandler : handlerList) {
            writerExcel(onWriterDataHandler);
        }
        return this;
    }

    /**
     * 生成excel
     * @param handler
     * @return
     */
    public WriteExcelUtil writerExcel(OnWriterDataHandler handler) {
        return this.writerExcel(handler,null);
    }

    /**
     * 生成excel
     *
     * @param handler
     */
    public WriteExcelUtil writerExcel(OnWriterDataHandler handler,Integer _sheetIndex) {
        long totalRows = 1L;
        long begin = System.currentTimeMillis();

        handler.getWorkbook(wb);//把workBook对象放出去

        String sheetName = handler.setSheetName();
        if (sheetName == null || sheetName.isEmpty()) {//如果没有sheet，则认为数据为空
            throw new ExcelException("sheetName cannot be null");
        }

        Sheet sheet;//sheet对象
        int sheetIndex;//sheet的索引
        if(_sheetIndex == null){
            //初始化表格
            sheet = wb.createSheet(sheetName);
            sheetIndex = wb.getSheetIndex(sheet);
        }else{
            sheet = wb.getSheetAt(_sheetIndex);
            sheetIndex = _sheetIndex;
        }

        log.debug("set excel columnWidth");
        List<Integer> setColumnWidth = handler.setColumnWidth();//以一个字符的1/265.95的宽度作为一个单位,3000就是11.7,就是约等于11个字符,这里说的都是半角
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

            log.debug("set rowHeight current row is {}",i);
            Integer rowHeight = handler.setRowHeight(sheetIndex,i);//设置行高
            if(rowHeight!=null && !rowHeight.equals(0)){
                row.setHeight(rowHeight.shortValue());
            }

            log.debug("get row data current row is {}",i);
            List<String> list = handler.rowDataHandler(sheetIndex, i);//每一行的数据
            if(list == null || list.size() == 0){
                continue;
            }

            log.debug("while cell current row is {}",i);
            Integer cellType;//单元格样式
            //循环一行里的所有列
            for (int j = 0; j < list.size(); j++) {
                Cell cell = row.createCell(j);

                String value = list.get(j);
                if(value == null){//没数据认为是空
                    value = ExcelConstant.getEmptyCellValue();
                }else{//去掉前后空格
                    value = value.trim();
                }

                log.debug("set cell data,current row is {},current cell is {}",i,j);
                cellType = handler.setCellType(sheetIndex,i,j);
                if(cellType == null){
                    cell.setCellType(Cell.CELL_TYPE_STRING);//设置所有数据都是string
                    if (!ExcelConstant.getEmptyCellValue().equals(value)) {//单元格赋值
                        cell.setCellValue(value);
                    }
                }else if(cellType.equals(Cell.CELL_TYPE_NUMERIC)){//数字
                    cell.setCellType(cellType);//单元格样式使用指定样式
                    if (!ExcelConstant.getEmptyCellValue().equals(value)) {//单元格赋值
                        cell.setCellValue(Double.parseDouble(value));
                    }
                }else if(cellType.equals(Cell.CELL_TYPE_FORMULA)){//公式
                    cell.setCellType(cellType);//单元格样式使用指定样式
                    if (!ExcelConstant.getEmptyCellValue().equals(value)) {//单元格赋值
                        cell.setCellFormula(value);
                    }
                }else{//其他
                    cell.setCellType(cellType);//单元格样式使用指定样式
                    if (!ExcelConstant.getEmptyCellValue().equals(value)) {//单元格赋值
                        cell.setCellValue(value);
                    }
                }


                log.debug("set cell style,current row is {},current cell is {}",i,j);
                CellStyle cellStyle1 = handler.setCellStyle(sheetIndex,i,j,wb);
                if (cellStyle1 != null) {
                    cell.setCellStyle(cellStyle1);
                }

            }
        }

        log.debug("excel drawing picture");
        List<PictureDto> pictureDtoList = handler.setPicture();
        if(pictureDtoList != null && pictureDtoList.size() != 0){
            //创建画图类
            Drawing drawingPatriarch = sheet.createDrawingPatriarch();

            for (PictureDto pictureDto : pictureDtoList) {
                int pictureIndex = wb.addPicture(pictureDto.getPictureData(), Workbook.PICTURE_TYPE_JPEG);

                //图片坐标
                ClientAnchor anchor = drawingPatriarch.createAnchor(pictureDto.getDx1(),pictureDto.getDy1(),pictureDto.getDx2(),pictureDto.getDy2(),pictureDto.getCol1(),pictureDto.getRow1(),pictureDto.getCol2(),pictureDto.getRow2());
                //Sets the anchor type （图片在单元格的位置）
                //0 = Move and size with Cells, 2 = Move but don't size with cells, 3 = Don't move or size with cells.
                anchor.setAnchorType(ClientAnchor.DONT_MOVE_AND_RESIZE);
                //添加图片
                drawingPatriarch.createPicture(anchor,pictureIndex );

            }
        }


        log.debug("excel merge cell");
        //合并单元格
        List<CellRangeAddressDto> cellRangeAddressDtoList = handler.addMergedRegions();
        if(cellRangeAddressDtoList != null && cellRangeAddressDtoList.size() != 0){
            for (CellRangeAddressDto cellRangeAddressDto : cellRangeAddressDtoList) {
                sheet.addMergedRegion(cellRangeAddressDto);

                if(cellRangeAddressDto.getBorderTop() != null){//上边框线条
                    RegionUtil.setBorderTop(cellRangeAddressDto.getBorderTop(),cellRangeAddressDto,sheet,wb);
                }
                if(cellRangeAddressDto.getBorderBottom() != null){//下边框线条
                    RegionUtil.setBorderBottom(cellRangeAddressDto.getBorderBottom(),cellRangeAddressDto,sheet,wb);
                }
                if(cellRangeAddressDto.getBorderLeft() != null){//左边框线条
                    RegionUtil.setBorderLeft(cellRangeAddressDto.getBorderLeft(),cellRangeAddressDto,sheet,wb);
                }
                if(cellRangeAddressDto.getBorderRight() != null){//右边框线条
                    RegionUtil.setBorderRight(cellRangeAddressDto.getBorderRight(),cellRangeAddressDto,sheet,wb);
                }
            }
        }

        log.info(String.format("Excel数据写入并处理完成,共写入数据：%s行,耗时：%s seconds.", totalRows, (System.currentTimeMillis() - begin) / 1000F));

        return this;
    }

    /**
     * 创建一个workBook
     *
     * @param type
     * @return
     */
    private Workbook createWorkbook(int type) {
        if (type == ExcelType.EXCEL2003){
            wb = new HSSFWorkbook();
        }else if(type == ExcelType.EXCEL2007ByBulkData){
            wb = new SXSSFWorkbook();
        }else{
            wb = new XSSFWorkbook();
        }
        return wb;
    }

}
