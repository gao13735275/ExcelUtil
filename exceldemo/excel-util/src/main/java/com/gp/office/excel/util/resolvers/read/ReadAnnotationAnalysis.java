package com.gp.office.excel.util.resolvers.read;

import com.gp.office.excel.util.annotation.ReadExcelModel;
import com.gp.office.excel.util.annotation.ReadExcelModelColumn;
import com.gp.office.excel.util.model.read.ReadExcelAnnotationDto;
import com.gp.office.excel.util.model.read.ReadExcelColumnAnnotationDto;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.*;

/**
 * @program: excelutil
 * @description:
 * @author: gaopo
 * @create: 2018-10-08 10:22
 **/
public class ReadAnnotationAnalysis {
    /**
     * 解析read系列的所有注解
     * @param classOfT
     * @return
     */
    public static ReadExcelAnnotationDto analysis(Class classOfT){

        Map<String, ReadExcelColumnAnnotationDto> readExcelColumnAnnotationDtoMap = new HashMap<>();
        List<ReadExcelColumnAnnotationDto> vagueReadExcelColumnAnnotationDtoList = new ArrayList<>();//模糊匹配的数据,只支持后模糊
        ReadExcelAnnotationDto readExcelAnnotationDto = new ReadExcelAnnotationDto();
        readExcelAnnotationDto.setReadExcelColumnAnnotationDtoMap(readExcelColumnAnnotationDtoMap);
        readExcelAnnotationDto.setVagueReadExcelColumnAnnotationDtoList(vagueReadExcelColumnAnnotationDtoList);

        // 判断在类上是否使用了自定的注解
        boolean isExist = classOfT.isAnnotationPresent(ReadExcelModel.class);
        if (isExist){
            // 从类上获得自定义的注解
            ReadExcelModel readExcelModel = (ReadExcelModel) classOfT.getAnnotation(ReadExcelModel.class);
            readExcelAnnotationDto.setParentColumn(readExcelModel.parentColumn());//设置是否需要读继承的父类的属性
        }else{
            readExcelAnnotationDto.setParentColumn(false);//不需要读
        }


        Field[] fields;
        //需要父的注解
        if(readExcelAnnotationDto.getParentColumn()){
            for (Class superclass=classOfT; superclass != Object.class; superclass = superclass.getSuperclass()) {//循环取父类
                fields =  superclass.getDeclaredFields();//读取当前类的属性的对象
                setReadExcelColumnAnnotationDto(readExcelColumnAnnotationDtoMap,fields);//复制属性数据
            }
        }else{
            fields =  classOfT.getDeclaredFields();//读取当前类的属性的对象
            setReadExcelColumnAnnotationDto(readExcelColumnAnnotationDtoMap,fields);//复制属性数据
        }

        //设置模糊匹配的列的数据
        Set<String> titleNames = readExcelColumnAnnotationDtoMap.keySet();
        for (String titleName : titleNames) {
            if(titleName.endsWith("%")){//有模糊匹配的通配符
                vagueReadExcelColumnAnnotationDtoList.add(readExcelColumnAnnotationDtoMap.get(titleName));
            }
        }

        return readExcelAnnotationDto;
    }

    /**
     * 保存属性数据
     * @param readExcelColumnAnnotationDtoMap
     * @param fields
     */
    private static void setReadExcelColumnAnnotationDto(Map<String,ReadExcelColumnAnnotationDto> readExcelColumnAnnotationDtoMap, Field[] fields){
        ReadExcelColumnAnnotationDto readExcelColumnAnnotationDto;
        String fieldName;
        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];
            fieldName = field.getName();

            //一个属性，保存一个属性对应的所有注解的值
            readExcelColumnAnnotationDto = new ReadExcelColumnAnnotationDto();
            readExcelColumnAnnotationDto.setField(field);//属性对象

            //获取属性上的指定类型的注解
            Annotation annotation = field.getAnnotation(ReadExcelModelColumn.class);
            if(annotation != null){
                ReadExcelModelColumn readExcelModelColumn = (ReadExcelModelColumn)annotation;
                readExcelColumnAnnotationDtoMap.put(readExcelModelColumn.titleName(),readExcelColumnAnnotationDto);
                readExcelColumnAnnotationDto.setTitleName(readExcelModelColumn.titleName());
                readExcelColumnAnnotationDto.setRemoveTitleNamePrefix(readExcelModelColumn.removeTitleNamePrefix());//要移除的抬头的前缀
                readExcelColumnAnnotationDto.setRequired(readExcelModelColumn.required());//是否必填
                readExcelColumnAnnotationDto.setOnlyOneRequired(readExcelModelColumn.onlyOneRequired());//是否必填,只要有一个不为空，认为必填
                readExcelColumnAnnotationDto.setValueProcessRule(readExcelModelColumn.valueProcessRule());
                readExcelColumnAnnotationDto.setDateFormat(readExcelModelColumn.dateFormat());
                readExcelColumnAnnotationDto.setReplaceToEmpty(readExcelModelColumn.replaceToEmpty());
                readExcelColumnAnnotationDto.setTrim(readExcelModelColumn.isTrim());
            }else{//没写注解，则赋值
                readExcelColumnAnnotationDtoMap.put(fieldName,readExcelColumnAnnotationDto);
                readExcelColumnAnnotationDto.setTitleName(fieldName);
                readExcelColumnAnnotationDto.setRequired(false);//值不必填
            }
        }
    }
}
