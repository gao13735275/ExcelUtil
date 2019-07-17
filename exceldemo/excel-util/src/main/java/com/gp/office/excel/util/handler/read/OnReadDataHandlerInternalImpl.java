package com.gp.office.excel.util.handler.read;


import com.gp.office.excel.util.constants.ExcelConstant;
import com.gp.office.excel.util.exception.ExcelException;
import com.gp.office.excel.util.model.read.ReadExcelAnnotationDto;
import com.gp.office.excel.util.model.read.ReadExcelColumnAnnotationDto;
import com.gp.office.excel.util.resolvers.read.ReadAnnotationAnalysis;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @program: excelutil
 * @description:
 * @author: gaopo
 * @create: 2018-10-09 14:40
 **/
public class OnReadDataHandlerInternalImpl<T> extends BaseOnReadDataHandlerInternal {
    //注解的数据
    ReadExcelAnnotationDto readExcelAnnotationDto;
    //抬头数据
    List<String> excelTitleName = null;
    //抬头行
    int titleIndex;
    //数据行
    int startRowDateIndex;
    List<T> list = new ArrayList<>();//读取的数据

    Class<T> classOfT;//输出的对象

    OnReadTitleInternalHandler onReadTitleInternalHandler;//列头监听

    /**
     * 初始化
     * @param _classOfT
     * @param _titleIndex
     * @param _startRowDateIndex
     */
    public OnReadDataHandlerInternalImpl(Class _classOfT, int _titleIndex, int _startRowDateIndex) {
        //把所有注解拿出来
        titleIndex = _titleIndex;
        startRowDateIndex = _startRowDateIndex;
        classOfT = _classOfT;
        readExcelAnnotationDto = ReadAnnotationAnalysis.analysis(classOfT);
    }

    public OnReadDataHandlerInternalImpl(Class _classOfT,int _titleIndex, int _startRowDateIndex,OnReadTitleInternalHandler _onReadTitleInternalHandler){
        this(_classOfT,_titleIndex,_startRowDateIndex);
        onReadTitleInternalHandler = _onReadTitleInternalHandler;
    }

    /**
     * 取组装好的数据
     * @return
     */
    public List<T> getDateList(){
        return list;
    }

    /**
     * 不能设置没赋值的构造函数
     */
    private OnReadDataHandlerInternalImpl(){}

    @Override
    protected void tableDataTop(int sheetIndex, int rowIndex, List<String> rowData) {
        if(sheetIndex != 0){//现在只支持一个sheet
            return;
        }
    }

    @Override
    protected void tableData(int sheetIndex, int rowIndex, List<String> rowData) {
        if(rowIndex == titleIndex){//指定那行的数据
            excelTitleName = rowData;
            if(onReadTitleInternalHandler != null){//输出列头的数据
                onReadTitleInternalHandler.titleNameList(rowData);
            }
        }

        if(rowIndex >= startRowDateIndex){//数据
            T model;

            try {// TODO: 2018/3/30 这里到时候要修改，改成能把错误传递出去
                model = classOfT.newInstance();
            } catch (InstantiationException e) {
                return;
            } catch (IllegalAccessException e) {
                return;
            }

            String titleName;
            String value;
            for (int i = 0; i < rowData.size(); i++) {//赋值
                titleName = excelTitleName.get(i);
                ReadExcelColumnAnnotationDto readExcelColumnAnnotationDto = readExcelAnnotationDto.getReadExcelColumnAnnotationDtoMap().get(titleName);
                if(readExcelColumnAnnotationDto == null){
                    for (ReadExcelColumnAnnotationDto excelColumnAnnotationBO : readExcelAnnotationDto.getVagueReadExcelColumnAnnotationDtoList()) {//看看模糊匹配的数据
                        if(titleName.startsWith(excelColumnAnnotationBO.getTitleName().substring(0,excelColumnAnnotationBO.getTitleName().length()-1))){
                            readExcelColumnAnnotationDto = excelColumnAnnotationBO;
                        }
                    }

                    if(readExcelColumnAnnotationDto == null){//不用保存
                        continue;
                    }
                }

                if(readExcelColumnAnnotationDto.getTrim()){//去掉空格
                    value = rowData.get(i).trim();//单元格的值
                }else{
                    value = rowData.get(i);//单元格的值
                }


                //如果是必填，值又是空的，则报错
                if(readExcelColumnAnnotationDto.isRequired() && readExcelColumnAnnotationDto.isRequired() && value.equals(ExcelConstant.getEmptyCellValue())){
                    StringBuilder stringBuilder = new StringBuilder();
                    stringBuilder.append("第 ");
                    stringBuilder.append(rowIndex+1);
                    stringBuilder.append(" 行 ");
                    stringBuilder.append(titleName);
                    stringBuilder.append(" 必填");
                    throw new ExcelException(stringBuilder.toString());//第x行xxx必填
                }

                setDate(value,model,readExcelColumnAnnotationDto,titleName);
            }

            for (ReadExcelColumnAnnotationDto readExcelColumnAnnotationDto : readExcelAnnotationDto.getVagueReadExcelColumnAnnotationDtoList()) {//模糊匹配必填项
                try {
                    if(readExcelColumnAnnotationDto.isOnlyOneRequired() && readExcelColumnAnnotationDto.getField().get(model) == null){
                        StringBuilder stringBuilder = new StringBuilder();
                        stringBuilder.append("第 ");
                        stringBuilder.append(rowIndex+1);
                        stringBuilder.append(" 行 ");
                        stringBuilder.append(readExcelColumnAnnotationDto.getTitleName(), 0, readExcelColumnAnnotationDto.getTitleName().length()-1);//提示的时候把百分号去掉
                        stringBuilder.append(" 至少填写一个");
                        throw new ExcelException(stringBuilder.toString());//第x行xxx必填
                    }
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            }

            list.add(model);
        }
    }

    @Override
    protected void tableDataDown(int sheetIndex, int rowIndex, List<String> rowData) {

    }

    @Override
    protected int tableDataTopRowCount(int sheetIndex) {
        return titleIndex;
    }

    @Override
    protected int tableDataDownRowCount(int sheetIndex) {
        return 0;
    }

    /**
     * 数据类型转化
     * @param cellData 数据值
     * @param model 反射出来的model
     */
    protected static void setDate(String cellData,Object model,ReadExcelColumnAnnotationDto readExcelColumnAnnotationDto,String titleName){
        if(cellData == null){//没有值就不写
            return;
        }

        Field field = readExcelColumnAnnotationDto.getField();
        // 设置类的私有字段属性可访问
        field.setAccessible(true);

        if(cellData.equals(ExcelConstant.getEmptyCellValue())){
            if (String.class == field.getType() && readExcelColumnAnnotationDto.getReplaceToEmpty()) {//没有值设置为""
                try {
                    field.set(model, "");
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
            return;
        }

        //正则验证有值，就写
        if(null != readExcelColumnAnnotationDto.getValueProcessRule() && !readExcelColumnAnnotationDto.getValueProcessRule().equals("")){
            Pattern p = Pattern.compile(readExcelColumnAnnotationDto.getValueProcessRule());
            Matcher m = p.matcher(cellData);
            while(m.find()){//没生效就不赋值
                cellData = m.group(1);
            }
        }

        try {
            if (String.class == field.getType()) {
                field.set(model, String.valueOf(cellData));
            } else if (BigDecimal.class == field.getType()) {
                cellData = cellData.indexOf("%") != -1 ? cellData.replace("%", "") : cellData;
                field.set(model, BigDecimal.valueOf(Double.valueOf(cellData)));
            } else if (java.util.Date.class == field.getType()) {
                SimpleDateFormat format =   new SimpleDateFormat( readExcelColumnAnnotationDto.getDateFormat().equals("")?"yyyy-MM-dd HH:mm:ss":readExcelColumnAnnotationDto.getDateFormat() );
                field.set(model,format.parse(cellData));
            } else if (java.sql.Date.class == field.getType()) {
                field.set(model,java.sql.Date.valueOf(cellData));
            }  else if ((Integer.TYPE == field.getType()) || (Integer.class == field.getType())) {
                field.set(model, Integer.parseInt(cellData));
            } else if ((Long.TYPE == field.getType()) || (Long.class == field.getType())) {
                field.set(model, Long.valueOf(cellData));
            } else if ((Float.TYPE == field.getType()) || (Float.class == field.getType())) {
                field.set(model, Float.valueOf(cellData));
            } else if ((Short.TYPE == field.getType()) || (Short.class == field.getType())) {
                field.set(model, Short.valueOf(cellData));
            } else if ((Double.TYPE == field.getType()) || (Double.class == field.getType())) {
                field.set(model, Double.valueOf(cellData));
            } else if (Character.TYPE == field.getType() || (Character.class == field.getType())) {
                if(cellData.length() > 0){
                    field.set(model, Character.valueOf(cellData.charAt(0)));
                }
            }else if(List.class == field.getType()){
                Object obj = field.get(model);
                List<String> list;
                if(obj == null){
                    list = new ArrayList<>();
                    field.set(model,list);
                }else{
                    list = (List<String>) obj;
                }
                list.add(cellData);

            }else if(Map.class == field.getType()){
                Object obj = field.get(model);
                Map<String,String> map;
                if(obj == null){
                    map = new LinkedHashMap<>();
                    field.set(model,map);
                }else{
                    map = (Map<String,String>) obj;
                }

                if(readExcelColumnAnnotationDto.getRemoveTitleNamePrefix() != null && !readExcelColumnAnnotationDto.getRemoveTitleNamePrefix().equals("")){//移除指定抬头前缀
                    //判断是否包含要移除的抬头
                    if(titleName.startsWith(readExcelColumnAnnotationDto.getRemoveTitleNamePrefix())){
                        String newTitleName = titleName.substring(titleName.indexOf(readExcelColumnAnnotationDto.getRemoveTitleNamePrefix()) + readExcelColumnAnnotationDto.getRemoveTitleNamePrefix().length());//移除掉抬头前缀
                        map.put(newTitleName,cellData);
                    }else{//没有匹配的，要移除的抬头
                        map.put(titleName,cellData);
                    }
                }else{//不用移除抬头前缀
                    map.put(titleName,cellData);
                }

            }
        } catch (IllegalAccessException e) {
        } catch (ParseException e) {
        }
    }
}
