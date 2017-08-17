package com.shure.utils.excel;

/**
 * Excel 内容存储对象
 */
public class ExcelDataVO {

	private int row;// Excel单元格行
	private int column;// Excel单元格列
	private String key;// 替换的关键字
	private String value;// 替换的文本
	
	private int firstRow; // 起始行
	private int endRow; // 结束行
	private int firstCol; // 起始列
	private int endCol; // 结束列
	
	private String[] textlist; // 下拉列表内容

	public ExcelDataVO(){}

	public int getRow() {
		return row;
	}

	public void setRow(int row) {
		this.row = row;
	}

	public int getColumn() {
		return column;
	}

	public void setColumn(int column) {
		this.column = column;
	}

	public String getKey() {
		return key;
	}

	public void setKey(String key) {
		this.key = key;
	}

	public String getValue() {
		return value;
	}

	public void setValue(String value) {
		this.value = value;
	}
	
	public int getFirstRow() {
		return firstRow;
	}

	public void setFirstRow(int firstRow) {
		this.firstRow = firstRow;
	}

	public int getEndRow() {
		return endRow;
	}

	public void setEndRow(int endRow) {
		this.endRow = endRow;
	}

	public int getFirstCol() {
		return firstCol;
	}

	public void setFirstCol(int firstCol) {
		this.firstCol = firstCol;
	}

	public int getEndCol() {
		return endCol;
	}

	public void setEndCol(int endCol) {
		this.endCol = endCol;
	}

	public String[] getTextlist() {
		return textlist;
	}

	public void setTextlist(String[] textlist) {
		this.textlist = textlist;
	}

}
