package com.gp.office.excel.util.handler.write;

import com.gp.office.excel.util.constants.ExcelConstant;
import com.gp.office.excel.util.model.write.CellRangeAddressDto;
import com.gp.office.excel.util.model.write.WriteExcelAnnotationDto;
import com.gp.office.excel.util.model.write.WriteExcelColumnAnnotationDto;
import com.gp.office.excel.util.resolvers.write.WriteAnnotationAnalysis;
import com.gp.office.excel.util.style.ExcelStyleUtil;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @program: excelutil
 * @description:
 * @author: gaopo
 * @create: 2018-10-08 17:29
 **/
public class OnWriterDataHandlerInternalImpl extends BaseOnWriterDataHandlerInternal {
    List<?> list;//数据
    List<String> currentRowData;//当前行组装好的数据
    Integer currentDataIndex = -2;//当前数据的索引
    WriteExcelAnnotationDto writeExcelAnnotationDto;
    List<Integer> columnWidthList = new ArrayList<>();
    List<String> titleList = new ArrayList<>();
    String sheetName;
    String banner;//最顶上大标题的内容


    public <T> OnWriterDataHandlerInternalImpl(String _sheetName,List<?> _list, Class<T> classOfT) {
        this(_sheetName,_sheetName,_list,classOfT);
    }

    /**
     * 初始化模版
     * @param _sheetName
     * @param _banner  大的标题
     * @param _list
     * @param classOfT
     * @param <T>
     */
    public <T> OnWriterDataHandlerInternalImpl(String _sheetName,String _banner,List<?> _list, Class<T> classOfT) {
        list = _list;
        writeExcelAnnotationDto = WriteAnnotationAnalysis.analysis(classOfT);
        sheetName = _sheetName;
        banner = _banner;
        init();
    }



    //不能设置没赋值的构造函数
    private OnWriterDataHandlerInternalImpl(){}

    /**
     * 设置列宽
     * @return
     */
    @Override
    public List<Integer> setColumnWidth() {
        return columnWidthList;
    }

    @Override
    public Integer setRowHeight(int sheetIndex, int rowIndex) {
        if(sheetIndex != 0){//现在只支持一个sheet
            return null;
        }

        if(rowIndex == 0){
            return 1000;
        }

        return null;
    }

    /**
     * 设置合并单元格的数据
     * @return
     */
    @Override
    public List<CellRangeAddressDto> addMergedRegions() {
        List<CellRangeAddressDto> cellRangeAddressDtoList = new ArrayList<>();
        cellRangeAddressDtoList.add(new CellRangeAddressDto(0,0,0,titleList.size()));
        return cellRangeAddressDtoList;
    }

    @Override
    public String setSheetName() {
        return sheetName;
    }

    /**
     * 设置样式
     * @param sheetIndex
     * @param rowIndex
     * @return
     */
    @Override
    public CellStyle setCellStyle(int sheetIndex, int rowIndex, int columnIndex, Workbook wb) {
        if(sheetIndex != 0){//现在只支持一个sheet
            return null;
        }

        if(rowIndex == 0){//顶部标题
            return ExcelStyleUtil.customHeadStyle(wb,writeExcelAnnotationDto.getFontSize());
        }

        if(rowIndex == tableDataTopRowCount(sheetIndex)){//标题
            WriteExcelColumnAnnotationDto writeExcelColumnAnnotationDto = writeExcelAnnotationDto.getWriterExcelAnnotationDtoList().get(columnIndex);
            if(writeExcelColumnAnnotationDto.isMark()){//列标记红色
                return ExcelStyleUtil.redTitleStyle(wb);
            }else{
                return ExcelStyleUtil.defaultTitleStyle(wb);
            }

        }

        return null;
    }

    @Override
    public List<String> tableDataTop(int sheetIndex, int rowIndex) {
        List<String> top = new ArrayList<>();
        top.add(banner);
        return top;
    }

    @Override
    public List<String> tableData(int sheetIndex, int rowIndex) {
        currentDataIndex++;//索引增加

        if(sheetIndex != 0){//现在只支持一个sheet
            return null;
        }

        if(rowIndex == tableDataTopRowCount(sheetIndex)){//标题
            return titleList;
        }

        if(rowIndex > tableDataTopRowCount(sheetIndex)){//内容
            currentRowData = new ArrayList<>();
            Object obj = list.get(currentDataIndex);

            for (WriteExcelColumnAnnotationDto writeExcelColumnAnnotationDto :  writeExcelAnnotationDto.getWriterExcelAnnotationDtoList()) {
                try {
                    Object value = writeExcelColumnAnnotationDto.getField().get(obj);//获取数据
                    currentRowData.add(getDate(value,writeExcelColumnAnnotationDto));//加工数据
                } catch (IllegalAccessException e) {//报错就写0，到时候这里要加日志记录
                    currentRowData.add(ExcelConstant.getEmptyCellValue());
                }
            }

            return currentRowData;
        }

        return null;
    }

    @Override
    public List<String> tableDataDown(int sheetIndex, int rowIndex) {
        return null;
    }

    @Override
    public int tableDataTopRowCount(int sheetIndex) {
        return 1;
    }

    @Override
    public int tableDataRowCount(int sheetIndex) {
        return list.size()+1;
    }

    @Override
    public int tableDataDownRowCount(int sheetIndex) {
        return 0;
    }

    /**
     * 初始化数据
     */
    private void init(){
        List<WriteExcelColumnAnnotationDto> writeExcelColumnAnnotationBOList = writeExcelAnnotationDto.getWriterExcelAnnotationDtoList();
        for (WriteExcelColumnAnnotationDto writeExcelColumnAnnotationBO : writeExcelColumnAnnotationBOList) {
            titleList.add(writeExcelColumnAnnotationBO.getTitleName());//抬头的列

            //每个列的宽度
            if(writeExcelColumnAnnotationBO.getWidth() == null || writeExcelColumnAnnotationBO.getWidth().equals(0)){//没设置列宽，默认是4000
                columnWidthList.add(ExcelConstant.getDefaultColumnWidth());
            }else{
                columnWidthList.add(writeExcelColumnAnnotationBO.getWidth());
            }
        }
    }

    /**
     * 赋值
     * @param data
     * @param writeExcelColumnAnnotationDto
     * @return
     */
    private static String getDate(Object data,WriteExcelColumnAnnotationDto writeExcelColumnAnnotationDto){
        if(data == null){
            return ExcelConstant.getEmptyCellValue();
        }
        String cellData;

        Class<?> classOf = writeExcelColumnAnnotationDto.getField().getType();


        if (java.util.Date.class == classOf) {
            SimpleDateFormat format =   new SimpleDateFormat( writeExcelColumnAnnotationDto.getDateFormat().equals("")?"yyyy-MM-dd HH:mm:ss":writeExcelColumnAnnotationDto.getDateFormat() );
            cellData = format.format(data);
        } else if (java.sql.Date.class == classOf) {
            SimpleDateFormat format =   new SimpleDateFormat( writeExcelColumnAnnotationDto.getDateFormat().equals("")?"yyyy-MM-dd HH:mm:ss":writeExcelColumnAnnotationDto.getDateFormat() );
            cellData = format.format(data);
        }else{
            cellData = data.toString();
        }

        //正则验证有值，就写
        if(null != writeExcelColumnAnnotationDto.getValueProcessRule() && !writeExcelColumnAnnotationDto.getValueProcessRule().equals("")){
            Pattern p = Pattern.compile(writeExcelColumnAnnotationDto.getValueProcessRule());
            Matcher m = p.matcher(cellData);
            while(m.find()){//没生效就不赋值
                cellData = m.replaceAll("");
            }
        }


        return cellData;
    }


}
