package com.gp.office.excel.util.resolvers.write;

import com.gp.office.excel.util.annotation.WriteExcelModel;
import com.gp.office.excel.util.annotation.WriteExcelModelColumn;
import com.gp.office.excel.util.model.write.WriteExcelAnnotationDto;
import com.gp.office.excel.util.model.write.WriteExcelColumnAnnotationDto;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

/**
 * @program: excelutil
 * @description:
 * @author: gaopo
 * @create: 2018-10-08 17:14
 **/
public class WriteAnnotationAnalysis {
    /**
     * 解析writer系列
     * @param classOfT
     * @return
     */
    public static WriteExcelAnnotationDto analysis(Class classOfT){
        ArrayList<WriteExcelColumnAnnotationDto> writeExcelColumnAnnotationDtoList = new ArrayList<>();
        WriteExcelAnnotationDto writeExcelAnnotationDto = new WriteExcelAnnotationDto();
        writeExcelAnnotationDto.setWriterExcelAnnotationDtoList(writeExcelColumnAnnotationDtoList);

        // 判断在类上是否使用了自定的注解
        boolean isExist = classOfT.isAnnotationPresent(WriteExcelModel.class);
        if (isExist){
            // 从类上获得自定义的注解
            WriteExcelModel writeExcelModel = (WriteExcelModel) classOfT.getAnnotation(WriteExcelModel.class);
            writeExcelAnnotationDto.setParentColumn(writeExcelModel.parentColumn());//设置是否需要读继承的父类的属性
            writeExcelAnnotationDto.setFontSize(writeExcelModel.fontSize());//字体大小
        }else{
            writeExcelAnnotationDto.setParentColumn(false);//不需要读
        }

        Field[] fields;
        //需要父的注解
        if(writeExcelAnnotationDto.getParentColumn()){
            for (Class superclass=classOfT; superclass != Object.class; superclass = superclass.getSuperclass()) {//循环取父类
                fields =  superclass.getDeclaredFields();//读取当前类的属性的对象
                setWriteExcelAnnotationDto(writeExcelColumnAnnotationDtoList,fields);//复制属性数据
            }
        }else{
            fields =  classOfT.getDeclaredFields();//读取当前类的属性的对象
            setWriteExcelAnnotationDto(writeExcelColumnAnnotationDtoList,fields);//复制属性数据
        }


        //根据sortNum排序
        writeExcelColumnAnnotationDtoList.sort(Comparator.comparing(WriteExcelColumnAnnotationDto::getSortNum));

        return writeExcelAnnotationDto;
    }

    private static void setWriteExcelAnnotationDto(List<WriteExcelColumnAnnotationDto> writeExcelColumnAnnotationDtoList, Field[] fields){
        WriteExcelColumnAnnotationDto writeExcelColumnAnnotationDto;
        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];

            //一个属性，保存一个属性对应的所有注解的值
            writeExcelColumnAnnotationDto = new WriteExcelColumnAnnotationDto();
            writeExcelColumnAnnotationDto.setField(field);//属性对象
            field.setAccessible(true);//设置所有属性可读


            //获取属性上的指定类型的注解
            Annotation annotation = field.getAnnotation(WriteExcelModelColumn.class);
            if (annotation != null) {//把写了注解的部分拿出来，没写不管
                WriteExcelModelColumn writeExcelModelColumn = (WriteExcelModelColumn) annotation;
                writeExcelColumnAnnotationDtoList.add(writeExcelColumnAnnotationDto);
                writeExcelColumnAnnotationDto.setTitleName(writeExcelModelColumn.titleName());
                writeExcelColumnAnnotationDto.setValueProcessRule(writeExcelModelColumn.valueProcessRule());
                writeExcelColumnAnnotationDto.setWidth(writeExcelModelColumn.width());
                writeExcelColumnAnnotationDto.setSortNum(writeExcelModelColumn.sortNum());
                writeExcelColumnAnnotationDto.setDateFormat(writeExcelModelColumn.dateFormat());
                writeExcelColumnAnnotationDto.setMark(writeExcelModelColumn.isMark());
            }
        }
    }
}
