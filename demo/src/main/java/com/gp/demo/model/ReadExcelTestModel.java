package com.gp.demo.model;

import com.gp.office.excel.util.annotation.ReadExcelModel;
import com.gp.office.excel.util.annotation.ReadExcelModelColumn;
import lombok.Data;

import java.time.LocalDateTime;
import java.util.Date;
import java.util.List;

/**
 * @program: excelHelper
 * @description:
 * @author: gaopo
 * @create: 2018-03-29 16:51
 **/
@Data
@ReadExcelModel(parentColumn = false)//如果改成false，则会给父类的属性赋值
public class ReadExcelTestModel extends ReadExcelTestParent{

    @ReadExcelModelColumn(titleName="公司名称")
    private String companyName;//公司名称

    private Character Salesman;//业务员

    @ReadExcelModelColumn(titleName = "订单时间",dateFormat="yyyy-MM-dd")
    private Date orderTime;//订单时间

    @ReadExcelModelColumn(titleName = "订单详细时间",dateFormat="yyyy-MM-dd")
    private LocalDateTime orderTimeDetail;//订单时间

    @ReadExcelModelColumn(titleName = "规格")
    private String specifications;//规格

    @ReadExcelModelColumn(titleName = "品种")
    private String varieties;//品种

    @ReadExcelModelColumn(titleName = "类型")
    private String productType;//产品类型

    @ReadExcelModelColumn(titleName = "产品图片%")
    private List<String> pictureUuid;
}
