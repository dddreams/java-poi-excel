package com.shure.utils.excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.util.List;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFName;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;


public class ExcelUtil {

	/**
	 * 生成 Excel 模板
	 * 
	 * @param sheets
	 *            sheet 对象
	 * @param sourceFilePath
	 *            文件路径
	 * @return
	 */
	public static boolean createExcelMod(List<ExcelSheetVO> sheets, String sourceFilePath, String outPath) {
		boolean res = true;
		try {
			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(sourceFilePath));
			HSSFWorkbook wb = new HSSFWorkbook(fs);

			// 删除默认 sheet
			for (int i = 0; i < wb.getNumberOfSheets(); i++) { // 获取每个Sheet表
				HSSFSheet defaultSheet = wb.getSheetAt(i);
				wb.removeSheetAt(wb.getSheetIndex(defaultSheet.getSheetName()));
			}

			if (sheets != null && sheets.size() > 0) {
				for (int i = 0, len = sheets.size(); i < len; i++) {
					ExcelSheetVO sheetVO = sheets.get(i);
					HSSFSheet sheet = wb.createSheet(sheetVO.getSheetName());
					sheet = setSheetData(wb, sheet, sheetVO);

					// 判断是否需要隐藏
					// if(sheetVO.getSheetName().indexOf("隐藏") >= 0){
					// wb.setSheetHidden(wb.getSheetIndex(sheetVO.getSheetName()),
					// false);
					// }
				}
			}
			// wb.setSheetHidden(0, true);

			// 输出文件
			FileOutputStream fileOut = new FileOutputStream(outPath);
			wb.write(fileOut);
			fileOut.close();

		} catch (Exception e) {
			res = false;
			e.printStackTrace();
		}
		return res;
	}

	/**
	 * 给 sheet 绑定数据
	 * 
	 * @param wb
	 * @param sheet
	 * @param sheetVO
	 * @return
	 */
	private static HSSFSheet setSheetData(HSSFWorkbook wb, HSSFSheet sheet, ExcelSheetVO sheetVO) throws Exception {

		// 生成表头
		HSSFRow header = sheet.createRow(0);
		HSSFCellStyle cellStyle = wb.createCellStyle();
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		String[] headers = sheetVO.getHeaders();
		if (headers != null && headers.length > 0) {
			for (int j = 0, len = headers.length; j < len; j++) {
				HSSFCell cell = header.createCell(j);
				if (headers[j] != null) {
					cell.setCellValue(headers[j]);
					// 设置字体
					HSSFFont font = wb.createFont();
					font.setFontName("仿宋_GB2312");
					font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);// 字体加粗
					font.setFontHeightInPoints((short) 14);
					cellStyle.setFont(font);
					cell.setCellStyle(cellStyle);

					// 设置列宽
					sheet.setColumnWidth(j, 256 * 15);
				}
			}
		}

		// 生成行
		for (int i = 1; i <= sheetVO.getMaxLine(); i++) {
			HSSFRow row = sheet.createRow(i);
			row.setHeight((short) 400);
		}

		// 设置数据
		List<ExcelDataVO> datas = sheetVO.getCells();
		if (datas != null && datas.size() > 0) {
			for (ExcelDataVO data : datas) {
				if (data.getValue() != null) {
					HSSFRow row = sheet.getRow(data.getRow());
					HSSFCell cell = row.createCell(data.getColumn());
					cell.setCellType(HSSFCell.CELL_TYPE_STRING);
					cell.setCellValue(data.getValue());
				}

				if (data.getTextlist() != null && data.getTextlist().length > 0) {
					sheet = setHSSFValidation(sheet, data.getTextlist(), data.getFirstRow(), data.getEndRow(),
							data.getFirstCol(), data.getEndCol());
				}
			}
		}

		return sheet;
	}

	/**
	 * 设置某些列的值只能输入预制的数据,显示下拉框.
	 * 
	 * @param sheet
	 *            要设置的sheet.
	 * @param textlist
	 *            下拉框显示的内容
	 * @param firstRow
	 *            开始行
	 * @param endRow
	 *            结束行
	 * @param firstCol
	 *            开始列
	 * @param endCol
	 *            结束列
	 * @return 设置好的sheet.
	 */
	public static HSSFSheet setHSSFValidation(HSSFSheet sheet, String[] textlist, int firstRow, int endRow,
			int firstCol, int endCol) {
		// 加载下拉列表内容
		DVConstraint constraint = DVConstraint.createExplicitListConstraint(textlist);
		// 设置数据有效性加载在哪个单元格上,四个参数分别是：起始行、终止行、起始列、终止列
		CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, firstCol, endCol);
		// 数据有效性对象
		HSSFDataValidation data_validation_list = new HSSFDataValidation(regions, constraint);
		sheet.addValidationData(data_validation_list);
		return sheet;
	}

	/**
	 * 设置单元格上提示
	 * 
	 * @param sheet
	 *            要设置的sheet.
	 * @param promptTitle
	 *            标题
	 * @param promptContent
	 *            内容
	 * @param firstRow
	 *            开始行
	 * @param endRow
	 *            结束行
	 * @param firstCol
	 *            开始列
	 * @param endCol
	 *            结束列
	 * @return 设置好的sheet.
	 */
	public static HSSFSheet setHSSFPrompt(HSSFSheet sheet, String promptTitle, String promptContent, int firstRow,
			int endRow, int firstCol, int endCol) {
		// 构造constraint对象
		DVConstraint constraint = DVConstraint.createCustomFormulaConstraint("BB1");
		// 四个参数分别是：起始行、终止行、起始列、终止列
		CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, firstCol, endCol);
		// 数据有效性对象
		HSSFDataValidation data_validation_view = new HSSFDataValidation(regions, constraint);
		data_validation_view.createPromptBox(promptTitle, promptContent);
		sheet.addValidationData(data_validation_view);
		return sheet;
	}

	/**
	 * 设置数据有效性（通过名称管理器级联相关）
	 * 
	 * @param name
	 * @param firstRow
	 * @param endRow
	 * @param firstCol
	 * @param endCol
	 * @return
	 */
	public static HSSFSheet setDataValidation(HSSFSheet sheet, String name, int firstRow, int endRow, int firstCol,
			int endCol) {
		// 加载下拉列表内容
		DVConstraint constraint = DVConstraint.createFormulaListConstraint(name);
		// 设置数据有效性加载在哪个单元格上。
		// 四个参数分别是：起始行、终止行、起始列、终止列
		CellRangeAddressList regions = new CellRangeAddressList((short) firstRow, (short) endRow, (short) firstCol,
				(short) endCol);
		// 数据有效性对象
		HSSFDataValidation data_validation = new HSSFDataValidation(regions, constraint);
		sheet.addValidationData(data_validation);
		return sheet;
	}

	/**
	 * 创建有效性名称
	 * 
	 * @param wb
	 * @param name
	 * @param expression
	 * @return
	 */
	public static HSSFName createName(HSSFWorkbook wb, String name, String expression) {
		HSSFName refer = wb.createName();
		refer.setRefersToFormula(expression);
		refer.setNameName(name);
		return refer;
	}

	/**
	 * 读取 excel 2003 and 2007
	 * @param inp
	 * @return
	 * @throws IOException
	 * @throws InvalidFormatException
	 */
	private static Workbook create(InputStream inp) throws IOException, InvalidFormatException {
		// If clearly doesn't do mark/reset, wrap up
		if (!inp.markSupported()) {
			inp = new PushbackInputStream(inp, 8);
		}

		if (POIFSFileSystem.hasPOIFSHeader(inp)) {
			return new HSSFWorkbook(inp);
		}
		if (POIXMLDocument.hasOOXMLHeader(inp)) {
			return new XSSFWorkbook(OPCPackage.open(inp));
		}
		throw new IllegalArgumentException("Your InputStream was neither an OLE2 stream, nor an OOXML stream");
	}

}
