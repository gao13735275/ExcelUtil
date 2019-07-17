package com.gp.office.excel.util;


import com.gp.office.excel.util.handler.read.OnReadDataHandlerInternalImpl;
import com.gp.office.excel.util.handler.read.OnReadTitleInternalHandler;
import com.gp.office.excel.util.resolvers.read.ReadExcelUtil;

import java.io.InputStream;
import java.util.List;

/**
 * @program: ourHouse
 * @description:
 * @author: gaopo
 * @create: 2019-06-20 15:25
 **/
public class ReadExcelApi {

    /**
     * 通过注解读excel
     * @param inputStream
     * @param classOfT 待读的数据类型
     * @param titleIndex 列抬头的位置
     * @param startRowDateIndex 数据起始位置
     * @return
     */
    public static <T> List<T> readExcelByAnnotation(InputStream inputStream, Class<T> classOfT, int titleIndex, int startRowDateIndex) {
        //初始化模板
        OnReadDataHandlerInternalImpl onReadDataHandlerInternalImpl = new OnReadDataHandlerInternalImpl<T>(classOfT,titleIndex,startRowDateIndex);
        //读excel
        ReadExcelUtil.readExcel(inputStream,onReadDataHandlerInternalImpl);
        //取数据
        return onReadDataHandlerInternalImpl.getDateList();
    }

    /**
     * 读取Excel数据
     * @param inputStream
     * @param classOfT 待读的数据类型
     * @param titleIndex 列抬头的位置
     * @param startRowDateIndex 数据起始位置
     * @param onReadTitleInternalHandler 获取列头的数据
     * @return
     */
    public static <T> List<T> readExcelByAnnotation(InputStream inputStream, Class<T> classOfT, int titleIndex, int startRowDateIndex, OnReadTitleInternalHandler onReadTitleInternalHandler) {
        //初始化模板
        OnReadDataHandlerInternalImpl onReadDataHandlerInternalImpl = new OnReadDataHandlerInternalImpl<T>(classOfT,titleIndex,startRowDateIndex,onReadTitleInternalHandler);
        //读excel
        ReadExcelUtil.readExcel(inputStream,onReadDataHandlerInternalImpl);
        //取数据
        return onReadDataHandlerInternalImpl.getDateList();
    }
}
