package com.gp.office.excel.util.model.write;

import org.apache.poi.ss.util.CellRangeAddress;

/**
 * @program: excelutil
 * @description:
 * @author: gaopo
 * @create: 2018-10-08 17:03
 **/
public class CellRangeAddressDto extends CellRangeAddress {

    public CellRangeAddressDto(int firstRow, int lastRow, int firstCol, int lastCol) {
        super(firstRow, lastRow, firstCol, lastCol);
    }
}
