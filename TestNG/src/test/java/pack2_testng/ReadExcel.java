package pack2_testng;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) throws IOException 
	{
		//Creating the file path and pass that file as the parameter
		File file = new File("C:\\Users\\Aparna.Venugopal\\eclipse-workspace\\TestNG_Proj2\\DataSource.xls");
		FileInputStream fis  = new FileInputStream(file);
		
		//Excel setup
		//HSSFWorkbook - Horrible Spread Sheet Format (.xls)
		//XSSFWorkbook - (.xlsx)
		HSSFWorkbook wb = new HSSFWorkbook(fis);
		
        //Create instance of excel 'sheet'
		//First sheet - index 0, second sheet index 1 like that
		HSSFSheet sheet = wb.getSheetAt(0);
		
		//Also can use sheet name
		//HSSFSheet sheet1 = wb.getSheet("Sheetname");
		
		//Get the number of rows in the sheet
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
		
		//Get the number of cells in a row 
		int cellCount = sheet.getRow(1).getLastCellNum();
		
		//Create array to store the values from the excel sheet
		String data[][] = new String[rowCount+1][cellCount];
		 
        for (int i=0; i<=rowCount;i++)
        {
        	for (int j=0;j<cellCount;j++)
        	{
        		data[i][j] = sheet.getRow(i).getCell(j).getStringCellValue();
                System.out.print(data[i][j] +' ');  

        	}
            System.out.println();

        }
       	}

    }
