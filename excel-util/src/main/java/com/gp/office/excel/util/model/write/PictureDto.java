package com.gp.office.excel.util.model.write;

/**
 * @program: ExcelUtil
 * @description: 一张图片
 * @author: gaopo
 * @create: 2019-08-07 14:41
 **/
public class PictureDto {
    /**
     * 2007版的基本坐标尺寸
     * 参数换算  比如单元格长度是5000，想图片长度写满单元格，则dx2=EMU_PER_PIXEL*5000
     */
    public static final int EMU_PER_POINT = 12700;

    /**
     * the x coordinate in EMU within the first cell.
     */
    int dx1;
    /**
     * the y coordinate in EMU within the first cell.
     */
    int dy1;
    /**
     * the x coordinate in EMU within the second cell.
     */
    int dx2;
    /**
     * the y coordinate in EMU within the second cell.
     */
    int dy2;
    /**
     * the column (0 based) of the first cell.
     */
    int col1;
    /**
     * the row (0 based) of the first cell.
     */
    int row1;
    /**
     * the column (0 based) of the second cell.
     */
    int col2;
    /**
     * the row (0 based) of the second cell.
     */
    int row2;
    /**
     * 图片数据
     */
    byte[] pictureData;


    public int getDx1() {
        return dx1;
    }

    public void setDx1(int dx1) {
        this.dx1 = dx1;
    }

    public int getDy1() {
        return dy1;
    }

    public void setDy1(int dy1) {
        this.dy1 = dy1;
    }

    public int getDx2() {
        return dx2;
    }

    public void setDx2(int dx2) {
        this.dx2 = dx2;
    }

    public int getDy2() {
        return dy2;
    }

    public void setDy2(int dy2) {
        this.dy2 = dy2;
    }

    public int getCol1() {
        return col1;
    }

    public void setCol1(int col1) {
        this.col1 = col1;
    }

    public int getRow1() {
        return row1;
    }

    public void setRow1(int row1) {
        this.row1 = row1;
    }

    public int getCol2() {
        return col2;
    }

    public void setCol2(int col2) {
        this.col2 = col2;
    }

    public int getRow2() {
        return row2;
    }

    public void setRow2(int row2) {
        this.row2 = row2;
    }

    public byte[] getPictureData() {
        return pictureData;
    }

    public void setPictureData(byte[] pictureData) {
        this.pictureData = pictureData;
    }
}
