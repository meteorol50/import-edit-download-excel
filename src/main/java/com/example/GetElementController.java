package com.example;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
public class GetElementController {
	
	String path = System.getProperty("user.dir") + "/excel/sample.xlsx";
	
	@RequestMapping("/")
	public String index() {
		return "index";
	}
	
	@RequestMapping("/getElement")
	public String excel() throws EncryptedDocumentException, IOException {
		Workbook excel = WorkbookFactory.create(new FileInputStream(path));
		Sheet sheet = excel.getSheet("Sheet1");
		System.out.println(sheet);
		Row row = sheet.getRow(0);
		System.out.println(row);
		Cell cell = row.getCell(0);
		System.out.println(cell.getStringCellValue());
		return "result";
	}

}
