package pack.net;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bouncycastle.asn1.cmc.GetCert;

public class Excelhotel {
	
	public static void main(String[] args) throws IOException {
		
		File file = new File("C:\\Users\\spark\\eclipse-workspace\\AsgarExcelhotel\\Excel\\Book1.xlsx");
		FileInputStream stream = new FileInputStream(file);
	Workbook workbook = new XSSFWorkbook(stream);
	Sheet sheet = workbook.getSheet("sheet");
       	Row row = sheet.getRow(2);
		Cell cell = row.getCell(2);
		String data = cell.getStringCellValue();
		if(data.equals("venky")) {
			cell.setCellValue("hii");
		}
		FileOutputStream out = new FileOutputStream(file);
		
	workbook.write(out);
	System.out.println("done");
	}

}
