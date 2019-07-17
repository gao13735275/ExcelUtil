package com.gp.office.excel.util.model.write;

import java.util.ArrayList;

/**
 * @program: excelutil
 * @description:
 * @author: gaopo
 * @create: 2018-10-08 17:15
 **/
public class WriteExcelAnnotationDto {
    /**
     * 写数据时，是只读子的还是把父的属性也读,true是读父属性，false是不读
     */
    private Boolean parentColumn;

    /**
     * 这个model是在excel写入到excel时，保存待输入的model里的列的注解的值
     */
    private ArrayList<WriteExcelColumnAnnotationDto> writerExcelAnnotationDtoList;

    /**
     * 标题的字体大小
     */
    private short fontSize;

    /*****************以下的get set方法*********************************************************/
    public Boolean getParentColumn() {
        return parentColumn;
    }

    public void setParentColumn(Boolean parentColumn) {
        this.parentColumn = parentColumn;
    }

    public ArrayList<WriteExcelColumnAnnotationDto> getWriterExcelAnnotationDtoList() {
        return writerExcelAnnotationDtoList;
    }

    public void setWriterExcelAnnotationDtoList(ArrayList<WriteExcelColumnAnnotationDto> writerExcelAnnotationDtoList) {
        this.writerExcelAnnotationDtoList = writerExcelAnnotationDtoList;
    }

    public short getFontSize() {
        return fontSize;
    }

    public void setFontSize(short fontSize) {
        this.fontSize = fontSize;
    }
}
