package com.gp.office.excel.util.model.write;

import java.lang.reflect.Field;

/**
 * @program: excelutil
 * @description:
 * @author: gaopo
 * @create: 2018-10-08 17:16
 **/
public class WriteExcelColumnAnnotationDto {
    /**
     * 属性对象
     */
    private Field field;
    /**
     * excel的列抬头的名字
     */
    private String titleName;
    /**
     * 日期的格式
     */
    private String dateFormat;
    /**
     * 列宽
     */
    private Integer width;
    /**
     * 顺序号，谁在前谁在后，默认0不排序
     */
    private Integer sortNum;
    /**
     * 获取的数据的规则，内容是正则
     */
    private String valueProcessRule;

    /**
     * 设置列头标红，true红,false黑色
     */
    private boolean isMark;

    /*****************以下的get set方法*********************************************************/

    public Field getField() {
        return field;
    }

    public void setField(Field field) {
        this.field = field;
    }

    public String getTitleName() {
        return titleName;
    }

    public void setTitleName(String titleName) {
        this.titleName = titleName;
    }

    public String getDateFormat() {
        return dateFormat;
    }

    public void setDateFormat(String dateFormat) {
        this.dateFormat = dateFormat;
    }

    public Integer getWidth() {
        return width;
    }

    public void setWidth(Integer width) {
        this.width = width;
    }

    public Integer getSortNum() {
        return sortNum;
    }

    public void setSortNum(Integer sortNum) {
        this.sortNum = sortNum;
    }

    public String getValueProcessRule() {
        return valueProcessRule;
    }

    public void setValueProcessRule(String valueProcessRule) {
        this.valueProcessRule = valueProcessRule;
    }

    public boolean isMark() {
        return isMark;
    }

    public void setMark(boolean mark) {
        isMark = mark;
    }
}
