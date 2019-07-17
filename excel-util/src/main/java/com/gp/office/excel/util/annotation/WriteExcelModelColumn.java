package com.gp.office.excel.util.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @program: excelutil
 * @description:
 * @author: gaopo
 * @create: 2018-10-08 13:52
 **/
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface WriteExcelModelColumn {
    /**
     * excel的列抬头的名字,必填
     * @return
     */
    String titleName();

    /**
     * 获取的数据的规则，内容是正则
     * @return
     */
    String valueProcessRule() default "";

    /**
     * 列宽,0代表不设置,用默认
     * 以一个字符的1/256的宽度作为一个单位,3000就是11.7,就是约等于11个字符,这里说的都是半角
     * @return
     */
    int width() default 0;

    /**
     * 顺序号，谁在前谁在后，默认0不排序
     * @return
     */
    int sortNum() default 0;

    /**
     * 日期格式设置,日期格式设置比正则优先级高
     * @return
     */
    String dateFormat() default "";

    /**
     * 是否列标红,true红，false黑
     * @return
     */
    boolean isMark() default false;

    // TODO: 2018/10/8 没做
    /**
     * 是否合计
     * @return
     */
    boolean isSum() default false;
}
