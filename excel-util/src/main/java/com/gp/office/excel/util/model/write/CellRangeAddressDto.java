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

    public static CellRangeAddressDto create(int firstRow, int lastRow, int firstCol, int lastCol){
        return new CellRangeAddressDto(firstRow, lastRow, firstCol, lastCol);
    }

    /**
     * 设置上边框的线条，系统默认是CellStyle.BORDER_NONE
     */
    private Short borderTop;
    /**
     * 设置下边框的线条，系统默认是CellStyle.BORDER_NONE
     */
    private Short borderBottom;
    /**
     * 设置左边框的线条，系统默认是CellStyle.BORDER_NONE
     */
    private Short borderLeft;
    /**
     * 设置右边框的线条，系统默认是CellStyle.BORDER_NONE
     */
    private Short borderRight;

    public Short getBorderTop() {
        return borderTop;
    }

    public CellRangeAddressDto setBorderTop(Short borderTop) {
        this.borderTop = borderTop;
        return this;
    }

    public Short getBorderBottom() {
        return borderBottom;
    }

    public CellRangeAddressDto setBorderBottom(Short borderBottom) {
        this.borderBottom = borderBottom;
        return this;
    }

    public Short getBorderLeft() {
        return borderLeft;
    }

    public CellRangeAddressDto setBorderLeft(Short borderLeft) {
        this.borderLeft = borderLeft;
        return this;
    }

    public Short getBorderRight() {
        return borderRight;
    }

    public CellRangeAddressDto setBorderRight(Short borderRight) {
        this.borderRight = borderRight;
        return this;
    }
}
