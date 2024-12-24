package DDT_Excel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingDataIntoExcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		FileOutputStream file = new FileOutputStream(System.getProperty("user.dir")+"\\testdata\\Automation.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Data");
		XSSFRow row1 = sheet.createRow(0);
			row1.createCell(0).setCellValue("Java");
			row1.createCell(1).setCellValue(21);
			row1.createCell(2).setCellValue("Auomation");
		XSSFRow row2 = sheet.createRow(1);
			row2.createCell(0).setCellValue("Python");
			row2.createCell(1).setCellValue(4);
			row2.createCell(2).setCellValue("Auomation");
		XSSFRow row3 = sheet.createRow(2);
			row3.createCell(0).setCellValue("Ruby");
			row3.createCell(1).setCellValue(3);
			row3.createCell(2).setCellValue("Auomation");
			
		workbook.write(file);
		workbook.close();
		file.close();
		
		System.out.println("File got created");

	}

}
