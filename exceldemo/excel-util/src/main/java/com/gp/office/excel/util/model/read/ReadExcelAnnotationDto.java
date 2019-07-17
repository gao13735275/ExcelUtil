package com.gp.office.excel.util.model.read;

import java.util.List;
import java.util.Map;

/**
 * @program: excelutil
 * @description:
 * @author: gaopo
 * @create: 2018-10-08 10:14
 **/
public class ReadExcelAnnotationDto {
    /**
     * 赋值属性时，是只赋值子的还是把父的属性也赋值,true是读父属性，false是不读
     */
    private Boolean parentColumn;

    /**
     * 这个model是在excel读取上传的excel时，保存待输出的model里的列的注解的值，其中key是注解里titleName的值，如果没写就是属性名
     */
    private Map<String,ReadExcelColumnAnnotationDto> readExcelColumnAnnotationDtoMap;

    /**
     * 模糊匹配的数据,只支持后模糊
     */
    private List<ReadExcelColumnAnnotationDto> vagueReadExcelColumnAnnotationDtoList = null;

    /*****************以下的get set方法*********************************************************/

    public Boolean getParentColumn() {
        return parentColumn;
    }

    public void setParentColumn(Boolean parentColumn) {
        this.parentColumn = parentColumn;
    }

    public Map<String, ReadExcelColumnAnnotationDto> getReadExcelColumnAnnotationDtoMap() {
        return readExcelColumnAnnotationDtoMap;
    }

    public void setReadExcelColumnAnnotationDtoMap(Map<String, ReadExcelColumnAnnotationDto> readExcelColumnAnnotationDtoMap) {
        this.readExcelColumnAnnotationDtoMap = readExcelColumnAnnotationDtoMap;
    }

    public List<ReadExcelColumnAnnotationDto> getVagueReadExcelColumnAnnotationDtoList() {
        return vagueReadExcelColumnAnnotationDtoList;
    }

    public void setVagueReadExcelColumnAnnotationDtoList(List<ReadExcelColumnAnnotationDto> vagueReadExcelColumnAnnotationDtoList) {
        this.vagueReadExcelColumnAnnotationDtoList = vagueReadExcelColumnAnnotationDtoList;
    }
}
