import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test1 {
	
	public ArrayList<String> excelData(String TestcaseName) throws IOException			// Test case name is the desired one
	{
		ArrayList<String> a = new ArrayList<String>();
		// 1.  Need to create object for XSSFWorkbook class
		// 2. Need to import excel file by using FileInputStrem method and access in object
		//	The excel file extensions should be xlsx
		// 3. Need to get the number of sheets
		// 4. To get the data sheet
		
		FileInputStream fis = new FileInputStream("C:\\Users\\Bablu\\OneDrive\\Desktop\\Medical Defence\\DemoData.xlsx");	//2
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook(fis);				//1
		
		int pages = workbook.getNumberOfSheets();			//3
		
		for(int i=0; i<pages; i++)					//4
		{
			if(workbook.getSheetName(i).equalsIgnoreCase("TestData"))
			{
				XSSFSheet sheet = workbook.getSheetAt(i);
				
				// 5. Identifying the Needed column by scanning the entire 1st row
				Iterator<Row> rows	=sheet.iterator();		// sheet is a collection of rows
				Row firstrow =rows.next();	// accessing the first row to check the data
				Iterator<Cell> ce =firstrow.cellIterator();	// Row is a combination of cells
				int k = 0;
				int column = 0;
				while (ce.hasNext())		// checking the row having next cell or not
				{
					Cell value =ce.next();	// Get the value of that cell
					if(value.getStringCellValue().equalsIgnoreCase("Data1"))		// Here test case name is the desired one
					{
						// desired column
						column =k;									// To find the index of the data c
					}
					k++;				
				}
				System.out.println(column);
				
				// Once column is identified then scan entire test case column  to identify desired test case row
				while(rows.hasNext())
				{
					Row r=	rows.next();
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase("Mahesh")) 
					{
						Iterator<Cell> c =r.cellIterator();
						while (c.hasNext())
						{
							a.add(c.next().getStringCellValue());	// To add the values in to array
						}
					}
					
					}
				}
		}
		return a;			// To return array list
	}

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		Test1 t = new Test1();
		ArrayList<String> data =t.excelData("Delete Profile");
		System.out.println(data.get(0));
		System.out.println(data.get(1));
		System.out.println(data.get(2));
		System.out.println(data.get(3));
		
	}

}
