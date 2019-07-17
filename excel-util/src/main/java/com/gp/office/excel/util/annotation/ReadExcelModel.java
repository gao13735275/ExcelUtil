package com.gp.office.excel.util.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @program: excelutil
 * @description: 在类上标记，表达是否读取父类的属性
 * @author: gaopo
 * @create: 2018-10-08 10:02
 **/
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ReadExcelModel {
    /**
     * 赋值属性时，是只赋值子的还是把父的属性也赋值,true是读父属性，false是不读
     * @return
     */
    boolean parentColumn() default true;
}
