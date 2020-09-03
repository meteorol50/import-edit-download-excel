package com.example;

import java.io.FileInputStream;
import java.io.OutputStream;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import com.example.model.ExcelView;

@Controller
@RequestMapping("/download")
public class DownloadController {
	
	private String path = System.getProperty("user.dir") + "/excel/sample.xlsx";
	
	@RequestMapping("/excel/import")
	public void excelImport(HttpServletResponse response) throws Exception {
		response.setHeader("Content-Disposition", "attachment; filename=\"sample.xlsx\"");
		response.setCharacterEncoding("UTF-8");
		Workbook workbook = WorkbookFactory.create(new FileInputStream(path));
		OutputStream out = response.getOutputStream();
		workbook.write(out);
		workbook.close();
	}
	
	@RequestMapping("/excel/create")
	public String excelCreate(HttpServletResponse response) throws Exception {
		response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		response.setHeader("Content-Disposition", "attachment; filename=\"sample.xlsx\"");
		response.setCharacterEncoding("UTF-8");
		Workbook workbook = new XSSFWorkbook();
		
		Sheet sheet = workbook.createSheet("aaa");
		createCell(sheet, 7, 7).setCellValue("abc");
		
		OutputStream out = response.getOutputStream();
		workbook.write(out);
		workbook.close();
		return null;  //これでもOK
	}
	
	@RequestMapping("/excel/edit")
	public void excelEdit(HttpServletResponse response) throws Exception {
		response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		response.setHeader("Content-Disposition", "attachment; filename=\"sample.xlsx\"");
		response.setCharacterEncoding("UTF-8");
		Workbook workbook = WorkbookFactory.create(new FileInputStream(path));
		
		Sheet sheet = workbook.getSheet("Sheet1");
		createCell(sheet, 7, 7).setCellValue(getCell(sheet, 0, 0).getStringCellValue());
		
		OutputStream out = response.getOutputStream();
		workbook.write(out);
		workbook.close();
	}
	
	@RequestMapping("/excel/edit2")
	public void excelEdit2(HttpServletResponse response) throws Exception {
		response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		response.setHeader("Content-Disposition", "attachment; filename=\"sample.xlsx\"");
		response.setCharacterEncoding("UTF-8");
		Workbook workbook = WorkbookFactory.create(new FileInputStream(path));
		
		Sheet sheet = workbook.getSheet("Sheet1");
		Row row = sheet.createRow(7);
		Cell cell = row.createCell(7);
		cell.setCellValue("テスト");
		
		OutputStream out = response.getOutputStream();
		workbook.write(out);
		workbook.close();
	}
	
	@RequestMapping("/excel2/edit")
	public ExcelView excel2(ExcelView excelView) throws Exception {
		return excelView;
	}
	
	/**
     * <p>
     * 引数で指定されたシートの、行番号、列番号で指定したセルを取得して返却する
     * <p>
     * 行番号、列番号は0から開始する
     * <p>
     * Excelテンプレートで該当のセルを操作していない場合、NullPointerExceptionになる
     * @param sheet シート
     * @param rowIndex 行番号
     * @param colIndex 列番号
     * @return セル
     */
    private Cell getCell(Sheet sheet, int rowIndex, int colIndex) {
        Row row = sheet.getRow(rowIndex);
        return row.getCell(colIndex);
    }
    
    private Cell createCell(Sheet sheet, int rowIndex, int colIndex) {
    	Row row = sheet.createRow(rowIndex);
    	return row.createCell(colIndex);
    }
    
}
