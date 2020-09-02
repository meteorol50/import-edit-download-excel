package com.example;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
public class GetContentController {
	
	String path = System.getProperty("user.dir") + "/excel/sample.xlsx";
	
	@RequestMapping("/getContent")
	public String excel() throws EncryptedDocumentException, IOException {
		Workbook excel = WorkbookFactory.create(new FileInputStream(path));
		//全セルを表示する
		for (Sheet sheet : excel ) {
			for (Row row : sheet) {
				for (Cell cell : row) {
					System.out.print(getCellValue(cell));
					System.out.print(" , ");
				}
				System.out.println();
			}
		}
		
		excel.close();
		return "result";
	}
	
	private static Object getCellValue(Cell cell) {
		switch (cell.getCellType()) {
			case STRING:
				return cell.getRichStringCellValue().getString();
			case NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					return cell.getDateCellValue();
				} else {
					return cell.getNumericCellValue();
				}
			case BOOLEAN:
				return cell.getBooleanCellValue();
			case FORMULA:
				return cell.getCellFormula();
			default:
				return null;
		}
	}
}
