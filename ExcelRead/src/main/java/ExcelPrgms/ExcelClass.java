package ExcelPrgms;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelClass {
	
	static FileInputStream f;
	static XSSFWorkbook w;
	static XSSFSheet sh;
	

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		String var;
		try {
			var = getStringData(1,1);
			int result=(int)getNumericData(1,0);
			
			System.out.println(var);
			System.out.println(result);
			
			
			
		} catch (IOException e) {
			
			e.printStackTrace();
		}
		
		
		
		
		
	}
	
	public static String getStringData(int a,int b) throws IOException {
		
		f=new FileInputStream("C:\\Users\\DELL\\Desktop\\excel1.xlsx");
		
		w=new XSSFWorkbook(f);
		sh=w.getSheet("Sheet1");
		
		Row r=sh.getRow(a);
		Cell c=r.getCell(b);
		
		return c.getStringCellValue();
		
		
		
	}
	
	
public static double getNumericData(int a,int b) throws IOException {
		
		f=new FileInputStream("C:\\Users\\DELL\\Desktop\\excel1.xlsx");
		
		w=new XSSFWorkbook(f);
		sh=w.getSheet("Sheet1");
		
		Row r=sh.getRow(a);
		Cell c=r.getCell(b);
		
		return c.getNumericCellValue();
		
		
		
	}

}
