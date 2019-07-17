package com.gp.demo.model;

import com.gp.office.excel.util.annotation.WriteExcelModelColumn;
import lombok.Data;

import java.util.Date;

/**
 * @program: excelHelper
 * @description:
 * @author: gaopo
 * @create: 2018-04-20 17:25
 **/
@Data
public class WriteExcelTestParent {
    @WriteExcelModelColumn(titleName = "开始时间",sortNum = 2)
    private Date startTime;//开始时间
}
