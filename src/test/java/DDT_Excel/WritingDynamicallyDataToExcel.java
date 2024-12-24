package DDT_Excel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingDynamicallyDataToExcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		FileOutputStream file = new FileOutputStream(System.getProperty("user.dir")+"\\testdata\\DynamicAutomationData.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet();
		Scanner sc = new Scanner(System.in);
		System.out.println("Enter number of Row");
		int noOrows = sc.nextInt();
		
		System.out.println("Enter number of cells");
		
		int noOfcells = sc.nextInt();
		
		
		for(int r=0;r<=noOrows;r++)
		{
			XSSFRow currentRow = sheet.createRow(r);
			for(int c=0;c<noOfcells;c++)
			{
				XSSFCell cell= currentRow.createCell(c);				
				cell.setCellValue(sc.next()); //user can enter any value that why sc.next() it will convert into string format
				
				
			}
			
		}
		workbook.write(file);
		workbook.close();
		file.close();
		
		System.out.println("Dynamic file is created .....");
	}

}
