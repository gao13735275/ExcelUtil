package com.gp.office.excel.util.style;


import org.apache.poi.ss.usermodel.*;

/**
 * @program: excelutil
 * @description:
 * @author: gaopo
 * @create: 2018-10-10 17:58
 **/
public class ExcelStyleUtil {
    /**
     * 默认的标题样式
     * @param wb
     * @return
     */
    public static CellStyle defaultHeadStyle(Workbook wb) {
        return customHeadStyle(wb,Short.valueOf("18"));
    }

    /**
     *
     * @param wb
     * @return
     */
    public static CellStyle customHeadStyle(Workbook wb,short fontSize) {
        CellStyle cellStyle = wb.createCellStyle();
        Font font = wb.createFont();
        // 设置表头样式
        //font.setBoldweight(Font.BOLDWEIGHT_BOLD); // 粗体
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font.setFontHeightInPoints(fontSize);
        cellStyle.setFont(font);
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);//水平居中
        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);//垂直居中
        cellStyle.setWrapText(true);//设置自动换行

        return cellStyle;
    }

    /**
     * 默认的列头的样式
     * @param wb
     * @return
     */
    public static CellStyle defaultTitleStyle(Workbook wb){
        CellStyle cellStyle = wb.createCellStyle();
        Font font = wb.createFont();
        //font.setBoldweight(Font.BOLDWEIGHT_BOLD); // 粗体
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font.setFontHeightInPoints(new Short("14"));
        cellStyle.setFont(font);
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);//水平居中
        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);//垂直居中
        return cellStyle;
    }

    /**
     * 红色的列头的样式
     * @param wb
     * @return
     */
    public static CellStyle redTitleStyle(Workbook wb){
        CellStyle cellStyle = wb.createCellStyle();
        Font font = wb.createFont();
        //font.setBoldweight(Font.BOLDWEIGHT_BOLD); // 粗体
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font.setFontHeightInPoints(new Short("14"));
        font.setColor(Font.COLOR_RED);//字体红色
        cellStyle.setFont(font);
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);//水平居中
        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);//垂直居中
        return cellStyle;
    }

    /**
     * 默认的合计的样式
     * @param wb
     * @return
     */
    public static CellStyle defaultTotalStyle(Workbook wb){
        CellStyle cellStyle = wb.createCellStyle();
        Font font = wb.createFont();
       // font.setBoldweight(Font.BOLDWEIGHT_BOLD); // 粗体
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font.setFontHeightInPoints(new Short("16"));
        font.setColor(Font.COLOR_RED);
        cellStyle.setFont(font);
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);//水平居中
        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);//垂直居中
        return cellStyle;
    }
}
