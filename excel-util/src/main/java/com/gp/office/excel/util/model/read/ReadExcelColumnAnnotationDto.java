package com.gp.office.excel.util.model.read;

import java.lang.reflect.Field;

/**
 * @program: excelutil
 * @description: 这个model是在excel读取上传的excel时，保存待输出的model里的列的注解的值
 * @author: gaopo
 * @create: 2018-10-08 10:16
 **/
public class ReadExcelColumnAnnotationDto {
    /**
     * 属性对象
     */
    private Field field;

    /**
     * 值是否必填
     */
    private boolean required;

    /**
     * 值是否必填，只要有一个填了，认为是符合条件的
     */
    private boolean onlyOneRequired;
    /**
     * 日期的格式
     */
    private String dateFormat;
    /**
     * excel的列抬头的名字
     */
    private String titleName;

    /**
     * 要移除的抬头名字的前缀
     */
    private String removeTitleNamePrefix;
    /**
     * 获取的数据的规则，内容是正则
     */
    private String valueProcessRule;
    /**
     * 返回值如果是null，写成空字符串，只适用于字符串类型
     */
    private Boolean replaceToEmpty;

    /**
     * 是否去掉头尾空格,false不去掉，true去掉
     */
    private Boolean isTrim;

    /*****************以下的get set方法*********************************************************/

    public Field getField() {
        return field;
    }

    public void setField(Field field) {
        this.field = field;
    }

    public boolean isRequired() {
        return required;
    }

    public void setRequired(boolean required) {
        this.required = required;
    }

    public boolean isOnlyOneRequired() {
        return onlyOneRequired;
    }

    public void setOnlyOneRequired(boolean onlyOneRequired) {
        this.onlyOneRequired = onlyOneRequired;
    }

    public String getDateFormat() {
        return dateFormat;
    }

    public void setDateFormat(String dateFormat) {
        this.dateFormat = dateFormat;
    }

    public String getTitleName() {
        return titleName;
    }

    public void setTitleName(String titleName) {
        this.titleName = titleName;
    }

    public String getRemoveTitleNamePrefix() {
        return removeTitleNamePrefix;
    }

    public void setRemoveTitleNamePrefix(String removeTitleNamePrefix) {
        this.removeTitleNamePrefix = removeTitleNamePrefix;
    }

    public String getValueProcessRule() {
        return valueProcessRule;
    }

    public void setValueProcessRule(String valueProcessRule) {
        this.valueProcessRule = valueProcessRule;
    }

    public Boolean getReplaceToEmpty() {
        return replaceToEmpty;
    }

    public void setReplaceToEmpty(Boolean replaceToEmpty) {
        this.replaceToEmpty = replaceToEmpty;
    }

    public Boolean getTrim() {
        return isTrim;
    }

    public void setTrim(Boolean trim) {
        isTrim = trim;
    }
}
