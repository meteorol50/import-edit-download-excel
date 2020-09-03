package com.example.model;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.web.servlet.view.document.AbstractXlsxView;

@Component
public class ExcelView extends AbstractXlsxView {
	
	String path = System.getProperty("user.dir") + "/excel/sample.xlsx";
	
	@Override
	protected Workbook createWorkbook(Map<String, Object> model, HttpServletRequest request) {
		try {
			return WorkbookFactory.create(new FileInputStream(path));
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
			return new XSSFWorkbook();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			return new XSSFWorkbook();
		} catch (IOException e) {
			e.printStackTrace();
			return new XSSFWorkbook();
		}
	}
	
	@Override
	protected void buildExcelDocument(Map<String, Object> model, Workbook workbook, HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		response.setHeader("Content-Disposition", "attachment; filename=\"samp.xlsx\"");
		response.setCharacterEncoding("UTF-8");
		Sheet sheet = workbook.getSheet("Sheet1");
		Row row = sheet.createRow(7);
		Cell cell = row.createCell(7);
		cell.setCellValue("テスト");
	}

}
