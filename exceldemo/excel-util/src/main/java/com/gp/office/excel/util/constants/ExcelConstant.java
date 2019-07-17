package com.gp.office.excel.util.constants;

/**
 * @program: excelutil
 * @description:
 * @author: gaopo
 * @create: 2018-10-08 10:06
 **/
public class ExcelConstant {
    /**
     * 默认以此值填充空单元格,可通过 setEmptyCellValue(string)改变其默认值。
     */
    private static String emptyCellValue = "EMPTY_CELL_VALUE";

    /**
     * 默认列宽
     */
    private static Integer defaultColumnWidth = 5000;

    /**
     * 设置没有字符串时的值，没有强迫症就不用写
     * @param emptyCellValue
     * @return
     */
    public void setEmptyCellValue(String emptyCellValue) {
        emptyCellValue = emptyCellValue;
    }

    public static String getEmptyCellValue() {
        return emptyCellValue;
    }

    public static Integer getDefaultColumnWidth() {
        return defaultColumnWidth;
    }

    public static void setDefaultColumnWidth(Integer defaultColumnWidth) {
        ExcelConstant.defaultColumnWidth = defaultColumnWidth;
    }
}
