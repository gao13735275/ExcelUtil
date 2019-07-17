package com.gp.demo.model;

import com.gp.office.excel.util.annotation.ReadExcelModelColumn;

import java.util.Date;


/**
 * @program: excelHelper
 * @description:
 * @author: gaopo
 * @create: 2018-04-20 14:36
 **/
@lombok.Data
public class ReadExcelTestParent {
    @ReadExcelModelColumn(titleName = "开始时间",dateFormat="yyyy-MM-dd")
    private Date startTime;//开始时间
}
