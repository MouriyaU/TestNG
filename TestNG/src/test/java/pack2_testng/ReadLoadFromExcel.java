package pack2_testng;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.testng.annotations.Test;

public class ReadLoadFromExcel {
public static void main(String[] args) throws IOException 
{
	//Creating the file path and pass that file as the parameter
	File file = new File("C:\\Users\\Aparna.Venugopal\\eclipse-workspace\\TestNG_Proj2\\Data_New.xls");
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




/*
package test;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
 
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.NumberFormat;
 
public class ExcelUtils {
   private static HSSFWorkbook workbook;
    private static HSSFSheet sheet;
    private static HSSFRow row;
    private static HSSFCell cell;
 
   public void setExcelFile(String excelFilePath,String sheetName) throws IOException {
       //Create an object of File class to open xls file
       File file =    new File("F:\\Selenium Material\\TestData.xls");

       //Create an object of FileInputStream class to read excel file
       FileInputStream inputStream = new FileInputStream(file);

       //creating workbook instance that refers to .xls file
       workbook=new HSSFWorkbook(inputStream);

       //creating a Sheet object
        sheet=workbook.getSheet(sheetName);
 
   }
 
    public String getCellData(int rowNumber,int cellNumber){
       //getting the cell value from rowNumber and cell Number
        cell =sheet.getRow(rowNumber).getCell(cellNumber);
        CellType cellType = cell.getCellType();

        //System.out.println(cellType);
        //returning the cell value as string
        switch(cellType)
        {
        case STRING:
            return cell.getStringCellValue();


        case NUMERIC:
            double retVal = cell.getNumericCellValue();
            System.out.println(retVal);
            NumberFormat nf = DecimalFormat.getInstance();
            nf.setMaximumFractionDigits(0);
            String str = nf.format(retVal);
            System.out.println(str);
            str=str.replace(",", "");
            System.out.println(str);

            return str;

        }
        return null;

    }
 
    public int getRowCountInSheet(){
       int rowcount = sheet.getLastRowNum()-sheet.getFirstRowNum();
       return rowcount;
    }
 
    public void setCellValue(int rowNum,int cellNum,String cellValue,String excelFilePath) throws IOException {
        //creating a new cell in row and setting value to it      
        sheet.getRow(rowNum).createCell(cellNum).setCellValue(cellValue);

        FileOutputStream outputStream = new FileOutputStream(excelFilePath);
        workbook.write(outputStream);
    }
}

*/