package com.svm.ExcelWritter;

import java.io.FileOutputStream;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWritter {

	public static Workbook getWorkbook(FileOutputStream out, String excelFilePath) {
		Workbook workbook = null;

		if (excelFilePath.endsWith("xlsx")) {
			workbook = new XSSFWorkbook();

		} else if (excelFilePath.endsWith("xls")) {
			workbook = new HSSFWorkbook();

		} else {
			throw new IllegalArgumentException("The specified file is not Excel file");
		}
		return workbook;
	}

	public static void write(Sheet sheet) {
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		Scanner scan =new Scanner(System.in);
		System.out.println("enter details to feed------"); 
		System.out.println("ID\t NAME\t LASTNAME");
		data.put("1", new Object[] { "ID", "NAME", "LASTNAME" });
		data.put("2", new Object[] { 1, "Amit", "Shukla" });
		data.put("3", new Object[] { 2, "Lokesh", "Gupta" });
		data.put("4", new Object[] { 3, "John", "Adwards" });
		data.put("5", new Object[] { 4, "Brian", "Schultz" });
		data.put("6", new Object[] { 5, "Shiv", "Bharadwaj" });
		data.put("7", new Object[] { 6.7, true, "D" });
		// Iterate over data and write to sheet
		
		Set<String> keyset = data.keySet();//Getting{1,2,3,4,5,6}
		
		int rownum = 0;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object[] objArr = data.get(key);
			
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof String)
					cell.setCellValue((String) obj);
				else if (obj instanceof Integer)
					cell.setCellValue((Integer) obj);
				else if (obj instanceof Boolean)
					cell.setCellValue((Boolean) obj);
				else if (obj instanceof Character)
					cell.setCellValue((Character) obj);
				else if (obj instanceof Double)
					cell.setCellValue((Double) obj);
			}
		}

	}

}
