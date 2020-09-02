package com.example;

import java.io.FileInputStream;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.web.servlet.view.document.AbstractXlsxView;

public class ExcelView extends AbstractXlsxView {
	
	String path = System.getProperty("user.dir") + "/excel/sample.xlsx";
	
	@Override
	protected void buildExcelDocument(Map<String, Object> model, Workbook workbook, HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		response.setHeader("Content-Disposition", "attachment; filename=\"sample.xlsx\"");
		response.setCharacterEncoding("UTF-8");
		workbook = null;
		workbook = WorkbookFactory.create(new FileInputStream(path));
		@SuppressWarnings("unused")
		Sheet sheet = workbook.getSheet("Sheet1");
	}

}
