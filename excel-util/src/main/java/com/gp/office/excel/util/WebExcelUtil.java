package com.gp.office.excel.util;


import com.gp.office.excel.util.constants.ExcelType;

import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;

/**
 * @program: ourHouse
 * @description:
 * @author: gaopo
 * @create: 2019-06-20 15:28
 **/
public class WebExcelUtil {
    private static final String CHAR_SET = "utf-8";

    /**
     * 获取excel用http返回时的head部分的设置
     *
     * @param type
     * @return
     */
    public static String getContentType(int type) {
        if (type == ExcelType.EXCEL2003) {
            return "application/vnd.ms-excel";
        } else {
            return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        }
    }

    /**
     * @param excelName 生成的excel文件名字
     * @param type      excel文件类型
     * @return
     * @throws Exception
     */
    public static String getHeader(String excelName, int type) {
        try {
            return "attachment; filename=" + URLEncoder.encode(String.format("%s%s", excelName, getExcelSuffix(type)), CHAR_SET);
        } catch (UnsupportedEncodingException e) {
            return null;
        }
    }

    /**
     * 获取excel的后缀名
     *
     * @param type
     * @return
     */
    public static String getExcelSuffix(int type) {
        if (type == ExcelType.EXCEL2003) {
            return ".xls";
        } else {
            return ".xlsx";
        }
    }
}
