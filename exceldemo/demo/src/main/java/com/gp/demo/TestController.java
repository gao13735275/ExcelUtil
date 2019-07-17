package com.gp.demo;

import com.gp.demo.model.ReadExcelTestModel;
import com.gp.demo.model.WriteExcelTestModel;
import com.gp.office.excel.util.ReadExcelApi;
import com.gp.office.excel.util.WebExcelUtil;
import com.gp.office.excel.util.WriteExcelApi;
import com.gp.office.excel.util.constants.ExcelType;
import com.gp.office.excel.util.exception.ExcelException;
import com.gp.office.excel.util.handler.read.OnReadTitleInternalHandler;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @program: exceldemo
 * @description:
 * @author: gaopo
 * @create: 2019-07-16 16:08
 **/
@RestController
public class TestController {

    /**
     * 导入
     * @param file
     */
    @PostMapping("/importExcel")
    public void importExcel(MultipartFile file){
        if(file == null){
            System.out.println("没有收到任何文件");
        }

        List<ReadExcelTestModel> readExcelTestModelList = null;
        try {
            readExcelTestModelList = ReadExcelApi.readExcelByAnnotation(file.getInputStream(), ReadExcelTestModel.class, 1, 2, new OnReadTitleInternalHandler() {
                @Override
                public void titleNameList(List<String> titleNameList) {//这里输出列头信息，用户列头的校验，不需要可以OnReadTitleInternalHandler这个不初始化，直接传空
                    System.out.println(titleNameList);
                }
            });
        } catch (IOException e) {
            System.out.println(e.toString());
        } catch (ExcelException e){
            System.out.println(e.toString());
        }
        System.out.println("转化数据，获取条数:" + readExcelTestModelList.size());

    }

    /**
     * 导出
     * @param response
     */
    @GetMapping("/exportExcel")
    public void exportExcel(HttpServletResponse response){
        try {
            String banner = "1.测试111 \n 2.测测测";

            response.setCharacterEncoding("utf-8");
            response.setContentType(WebExcelUtil.getContentType(ExcelType.EXCEL2007ByBulkData));
            response.setHeader("Content-disposition",WebExcelUtil.getHeader("excel的名字",ExcelType.EXCEL2007ByBulkData));

            WriteExcelApi.writerExcelByAnnotation(response.getOutputStream(),"sheet的名字",banner,this.getlist(),WriteExcelTestModel.class);

        } catch (IOException e) {
            System.out.println(e.toString());
        }


    }

    /**
     * 导出的数据
     * @return
     */
    private List<WriteExcelTestModel> getlist(){
        List<WriteExcelTestModel> list = new ArrayList<>();

        WriteExcelTestModel writeExcelTestModel1 = new WriteExcelTestModel();
        writeExcelTestModel1.setBatchNumber("批号221");
        writeExcelTestModel1.setCompanyName("公司名称1");
        writeExcelTestModel1.setOrderTime(new Date());
        writeExcelTestModel1.setProductType("产品类型1");
        writeExcelTestModel1.setSalesman("业务员1");
        writeExcelTestModel1.setSpecifications("规格1");
        writeExcelTestModel1.setVarieties("品种1");
        writeExcelTestModel1.setStartTime(new Date());
        writeExcelTestModel1.setEndTime(new java.sql.Date(new Date().getTime()));//结束时间
        list.add(writeExcelTestModel1);

        return list;
    }
}
