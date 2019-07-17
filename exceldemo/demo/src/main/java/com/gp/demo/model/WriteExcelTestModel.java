package com.gp.demo.model;


import com.gp.office.excel.util.annotation.WriteExcelModel;
import com.gp.office.excel.util.annotation.WriteExcelModelColumn;
import lombok.Data;

import java.util.Date;

/**
 * @program: excelHelper
 * @description:
 * @author: gaopo
 * @create: 2018-04-18 13:22
 **/
@Data
@WriteExcelModel(fontSize = 9)//修改标题大小，没事别赋值
public class WriteExcelTestModel extends WriteExcelTestParent{

    @WriteExcelModelColumn(titleName="公司名称",sortNum = 3)
    private String companyName;//公司名称

    private String Salesman;//业务员

    @WriteExcelModelColumn(titleName = "订单时间")
    private Date orderTime;//订单时间

    @WriteExcelModelColumn(titleName = "结束时间时间",dateFormat="yyyy-MM-dd")
    private java.sql.Date endTime;//结束时间

    @WriteExcelModelColumn(titleName = "批号",isMark = true)
    private String batchNumber;//批号

    @WriteExcelModelColumn(titleName = "规格")
    private String specifications;//规格

    @WriteExcelModelColumn(titleName = "品种")
    private String varieties;//品种

    @WriteExcelModelColumn(titleName = "类型")
    private String productType;//产品类型
}
