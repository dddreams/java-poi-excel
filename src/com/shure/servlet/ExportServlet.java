package com.shure.servlet;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.shure.utils.excel.ExcelSheetVO;
import com.shure.utils.excel.ExcelUtil;

/**
 * Servlet implementation class ExportServlet
 */
@WebServlet("/ExportServlet")
public class ExportServlet extends HttpServlet {
	private static final long serialVersionUID = 1L;

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		String docsPath = request.getSession().getServletContext().getRealPath("docs");
		String filePath = docsPath + "\\template.xls";
		try{
			String[] sheets = new String[]{"图书类别", "图书信息"};
			Map<Integer, String[]> headMap = new HashMap<Integer, String[]>();
			headMap.put(0, new String[]{"序号","图书名称","图书类别","作者","出版社","出版时间","价格"});
			headMap.put(1, new String[]{"序号","类型名称"});
			
			List<ExcelSheetVO> sheetVos = new ArrayList<ExcelSheetVO>();
			for(int i = 0, len = sheets.length; i < len; i++){
				ExcelSheetVO sheet = new ExcelSheetVO();
				sheet.setHeaders(headMap.get(i));
				sheet.setSheetName(sheets[i]);
				sheet.setMaxLine(1);
				sheetVos.add(sheet);
			}
			String fileName = "图书模板";
			boolean  b = ExcelUtil.createExcelMod(sheetVos, filePath, fileName + filePath.substring(filePath.lastIndexOf("."), filePath.length()));
			if(b)this.download(filePath, fileName, response);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		doGet(request, response);
	}
	
	/**
	 * 下载文件
	 * @param filePath
	 * @throws IOException
	 */
	private void download(String filePath, String fileName, HttpServletResponse response) throws IOException{
		File file = new File(filePath);
		fileName = fileName + file.getName().substring(file.getName().lastIndexOf("."), file.getName().length()); 
		// 以流的形式下载文件。
		InputStream fis = new BufferedInputStream(new FileInputStream(filePath));
		byte[] buffer = new byte[fis.available()];
		fis.read(buffer);
		fis.close();
		// 清空response
		response.reset();
		// 设置response的Header
		String fileNameNew = URLEncoder.encode(fileName, "UTF-8");
		fileNameNew = new String(fileName.getBytes("gb2312"), "ISO8859-1");
		response.addHeader("Content-Disposition", "attachment;filename=" + fileNameNew);
		response.addHeader("Content-Length", "" + file.length());
		OutputStream toClient = new BufferedOutputStream(response.getOutputStream());
		response.setContentType("application/vnd.ms-excel;charset=gb2312");
		toClient.write(buffer);
		toClient.flush();
		toClient.close();
	}

}
