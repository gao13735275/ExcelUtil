package com.gp.office.excel.util.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @program: excelutil
 * @description: 在类上标记，表达是否读取父类的属性
 * @author: gaopo
 * @create: 2018-10-08 13:51
 **/
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface WriteExcelModel {
    /**
     * 写数据时，是只读子的还是把父的属性也读,true是读父属性，false是不读
     * @return
     */
    boolean parentColumn() default true;

    /**
     * 最顶上第一行标题的字体大小
     * @return
     */
    short fontSize() default 18;
}
