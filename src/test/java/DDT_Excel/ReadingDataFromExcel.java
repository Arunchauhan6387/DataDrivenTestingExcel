package DDT_Excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import io.opentelemetry.exporter.logging.SystemOutLogRecordExporter;

public class ReadingDataFromExcel {

	public static void main(String[] args) throws IOException {
		
		//Reading data from excel
		FileInputStream file = new FileInputStream(System.getProperty("user.dir")+"\\testdata\\Data.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		//XSSFSheet sheet1 = workbook.getSheetAt(0);
		
		int totalRow = sheet.getLastRowNum();
		int totalCell = sheet.getRow(1).getLastCellNum();
		
		System.out.println("number of Rows "+totalRow);
		System.out.println("number of Cells "+totalCell);
		
		for(int r =0;r<=totalRow;r++) 
		{
			XSSFRow  currentRow =sheet.getRow(r);
			for(int c=0;c<totalCell;c++)
			{
				XSSFCell cell=currentRow.getCell(c);
				
				System.out.print(cell.toString()+"\t");
			}
			System.out.println();
		}
		workbook.close();
		file.close();

	}

}
