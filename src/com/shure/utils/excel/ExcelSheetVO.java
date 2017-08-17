package com.shure.utils.excel;

import java.util.List;

/**
 * Excel sheet 对象
 */
public class ExcelSheetVO {

    private String sheetName;
    private String[] headers;
    private List<ExcelDataVO> cells;
    private int maxLine;

    public ExcelSheetVO(){}

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public String[] getHeaders() {
        return headers;
    }

    public void setHeaders(String[] headers) {
        this.headers = headers;
    }

    public List<ExcelDataVO> getCells() {
        return cells;
    }

    public void setCells(List<ExcelDataVO> cells) {
        this.cells = cells;
    }

    public int getMaxLine() {
        return maxLine;
    }

    public void setMaxLine(int maxLine) {
        this.maxLine = maxLine;
    }
}
