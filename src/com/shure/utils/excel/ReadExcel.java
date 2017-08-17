package com.shure.utils.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	/**
	 * 对外提供读取excel 的方法
	 */
	public static List<List<Object>> readExcel(File file, int index) throws IOException {
		String fileName = file.getName();
		String extension = fileName.lastIndexOf(".") == -1 ? "" : fileName.substring(fileName.lastIndexOf(".") + 1);
		if ("xls".equals(extension)) {
			return read2003Excel(file, index);
		} else if ("xlsx".equals(extension)) {
			return read2007Excel(file, index);
		} else {
			throw new IOException("不支持的文件类型");
		}
	}

	/**
	 * 读取 office 2003 excel
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	private static List<List<Object>> read2003Excel(File file, int index) throws IOException {
		List<List<Object>> list = new LinkedList<List<Object>>();
		HSSFWorkbook hwb = new HSSFWorkbook(new FileInputStream(file));
		HSSFSheet sheet = hwb.getSheetAt(index);
		Object value = null;
		HSSFRow row = null;
		HSSFCell cell = null;

		DecimalFormat df = new DecimalFormat("0");// 格式化 number String
		// 字符
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");// 格式化日期字符串
		DecimalFormat nf = new DecimalFormat("0.00");// 格式化数字

		int maxCellNum = 0;
		for (int i = 0, len = sheet.getLastRowNum(); i <= len; i++) {
			row = sheet.getRow(i);
			if (row == null || isEmptyRow(row)) {
				continue;
			}
			if (i == 0) {
				maxCellNum = row.getLastCellNum();
			}
			List<Object> linked = new LinkedList<Object>();
			for (int j = row.getFirstCellNum(); j < maxCellNum; j++) {
				boolean isBlock = false;
				cell = row.getCell(j);
				if (cell == null) {
					linked.add("");
					continue;
				}

				switch (cell.getCellType()) {
				case XSSFCell.CELL_TYPE_STRING:
					value = cell.getStringCellValue();
					break;
				case XSSFCell.CELL_TYPE_NUMERIC:
					if ("@".equals(cell.getCellStyle().getDataFormatString())) {
						value = df.format(cell.getNumericCellValue());
					} else if ("General".equals(cell.getCellStyle().getDataFormatString())) {
						value = nf.format(cell.getNumericCellValue());
					} else {
						value = sdf.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()));
					}
					break;
				case XSSFCell.CELL_TYPE_BOOLEAN:
					value = cell.getBooleanCellValue();
					break;
				case XSSFCell.CELL_TYPE_BLANK:
					value = "";
					break;
				default:
					value = cell.toString();
				}
				
				linked.add(value);
			}
			list.add(linked);
		}
		return list;
	}

	/**
	 * 读取Office 2007 excel
	 */
	private static List<List<Object>> read2007Excel(File file, int index) throws IOException {
		List<List<Object>> list = new LinkedList<List<Object>>();
		// 构造 XSSFWorkbook 对象，strPath 传入文件路径
		XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(file));
		// 读取第一章表格内容
		XSSFSheet sheet = xwb.getSheetAt(index);
		Object value = null;
		XSSFRow row = null;
		XSSFCell cell = null;

		DecimalFormat df = new DecimalFormat("0");// 格式化 number String
		// 字符
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");// 格式化日期字符串
		DecimalFormat nf = new DecimalFormat("0.00");// 格式化数字
		int maxCellNum = 0;
		for (int i = 0, len = sheet.getLastRowNum(); i <= len; i++) {
			row = sheet.getRow(i);
			if (row == null || isEmptyRow(row)) {
				continue;
			}
			if (i == 0) {
				maxCellNum = row.getLastCellNum();
			}
			List<Object> linked = new LinkedList<Object>();
			for (int j = row.getFirstCellNum(); j < maxCellNum; j++) {
				cell = row.getCell(j);
				if (cell == null) {
					linked.add("");
					continue;
				}
				switch (cell.getCellType()) {
					case XSSFCell.CELL_TYPE_STRING:
						value = cell.getStringCellValue();
						break;
					case XSSFCell.CELL_TYPE_NUMERIC:
						if ("@".equals(cell.getCellStyle().getDataFormatString())) {
							value = df.format(cell.getNumericCellValue());
						} else if ("General".equals(cell.getCellStyle().getDataFormatString())) {
							value = nf.format(cell.getNumericCellValue());
						} else {
							value = sdf.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()));
						}
						break;
					case XSSFCell.CELL_TYPE_BOOLEAN:
						value = cell.getBooleanCellValue();
						break;
					case XSSFCell.CELL_TYPE_BLANK:
						value = "";
						break;
					default:
						value = cell.toString();
				}
				linked.add(value);
			}
			list.add(linked);
		}
		return list;
	}

	/**
	 * 判断是否为空行
	 *
	 * @param row 行对象
	 * @return true 空 false 非空
	 */
	public static boolean isEmptyRow(Row row) {
		if (row == null) {
			return true;
		}
		boolean result = true;
		for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
			Cell cell = row.getCell(i, HSSFRow.RETURN_BLANK_AS_NULL);
			String value = "";
			if (cell != null) {
				switch (cell.getCellType()) {
					case Cell.CELL_TYPE_STRING:
						value = cell.getStringCellValue();
						break;
					case Cell.CELL_TYPE_NUMERIC:
						value = String.valueOf((int) cell.getNumericCellValue());
						break;
					case Cell.CELL_TYPE_BOOLEAN:
						value = String.valueOf(cell.getBooleanCellValue());
						break;
					case Cell.CELL_TYPE_FORMULA:
						value = String.valueOf(cell.getCellFormula());
						break;
					default:
						break;
				}
				if (value != null && value.trim() != "" && "".equals(value.trim())) {
					result = false;
					break;
				}
			}
		}
		return result;
	}

	public static void main(String[] args) {

	}



}
