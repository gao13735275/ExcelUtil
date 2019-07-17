package com.gp.office.excel.util.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @program: excelutil
 * @description: bean的列注解
 * @author: gaopo
 * @create: 2018-10-08 10:04
 **/
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ReadExcelModelColumn {
    /**
     * excel的列抬头的名字
     * 支持尾匹配，头一样，尾模糊，格式是xxx%，模糊匹配必须是List<String>或Map<String,String>
     * Map<String,String>,存储的map真实类型是LinkedHashMap,是有序的,key是titleName,value是值，没有值的不会保存在map里
     * List<String>，存储的list真实类型是arrayList
     * @return
     */
    String titleName();

    /**
     * 一般与titleName的模糊匹配联合使用,在map返回方式中，移除一部分指定的名称的前缀
     * 本注解只是在最后放入map时，才过滤指定前缀，在这之前的判断，都以excel原始抬头为准
     * @return
     */
    String removeTitleNamePrefix() default "";
    /**
     * 值是否必填，默认是选填
     * 当返回值是list或map时，必须全部列都有值，才认为有值
     * @return
     */
    boolean required() default false;

    /**
     * 只要有一个是填了，就认为填了
     * 只有返回值是list或map时，该注解才生效，表达只要这些列，只要有任意一列是有值的，即认为有值
     * @return
     */
    boolean onlyOneRequired() default false;

    // TODO: 2019/4/4 正则数据没验证
    /**
     * 获取的数据的规则，内容是正则
     * @return
     */
    String valueProcessRule() default "";

    /**
     * 日期格式设置,正则优先级比日期格式高
     * @return
     */
    String dateFormat() default "";

    /**
     * 返回值如果是null，写成空字符串，只适用于字符串类型
     * @return
     */
    boolean replaceToEmpty() default false;

    /**
     * 是否去掉头尾空格,false不去掉，true去掉
     * @return
     */
    boolean isTrim() default false;
}
