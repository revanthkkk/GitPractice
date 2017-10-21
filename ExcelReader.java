package adac_UtilityFiles;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ExcelReader {

	public  File f;
	public  FileInputStream fis;
	public  FileOutputStream fout;
	public  String filepath;
	public  Workbook workbook;
	public  Sheet sheet;
	public  XSSFWorkbook writeworkbook;
	
	
	
	public ExcelReader(String filePath){
		
		filepath=filePath;
	    f=new File(filepath);
	    try{
	    fis=new FileInputStream(f);
		workbook=WorkbookFactory.create(fis);
	    }catch(InvalidFormatException |IOException e){
	    	
	    	e.printStackTrace();
	    	System.out.println("Error accured in initializing the Excel reader Object");
	    }
	}
	
	//main fuc for debugging
		public static void main(String[] args){
		
			//ExcelReader excel=new ExcelReader("C:\\NewTestPath\\DataDrivenFramework\\src\\TestData\\TestData.xlsx");
			
		    //System.out.println(excel.rowCount(Config.BOOKHOTELSHEETNAME));
		   // System.out.println(excel.readExcelData(Config.BOOKHOTELSHEETNAME, 2, 1));
			
		}//test main func close
		
	
	public String readExcelData(String sheetName,int rowNo,int colNo){
		
		String cellValue=null;
		
		try{
		fis=new FileInputStream(f);
		workbook=WorkbookFactory.create(fis);
		sheet=workbook.getSheet(sheetName);
		Row row=sheet.getRow(rowNo-1);
		Cell cell=row.getCell(colNo);
		cellValue=cell.getStringCellValue();
		fis.close();
		}catch(InvalidFormatException |IOException e){
	    	
	    	e.printStackTrace();
	    	System.out.println("Error accured while reading data from the Cell");
	    }
		
		return cellValue;
		
	}
	
	public int rowCount(String sheetName){
		int rowCount=0;
		
		try{
		fis=new FileInputStream(f);
		workbook=WorkbookFactory.create(fis);
		sheet=workbook.getSheet(sheetName);
		rowCount=sheet.getLastRowNum();
		fis.close();
		}catch(InvalidFormatException |IOException e){
	    	
	    	e.printStackTrace();
	    	System.out.println("Error accured while getting rowCount from the SheetName provided");
	    }
		
		return (rowCount+1);
	}

  public int columnCount(String sheetName,int rowNo){
	    int colCount=0;
	    try{
	    fis=new FileInputStream(f);
		workbook=WorkbookFactory.create(fis);
		sheet=workbook.getSheet(sheetName);
		Row row=sheet.getRow(rowNo-1);
		colCount=row.getLastCellNum();
		fis.close();
	    }catch(InvalidFormatException |IOException e){
	    	
	    	e.printStackTrace();
	    	System.out.println("Error accured while getting ColumnCount from the SheetName and RowNo provided");
	    }
		return colCount;
	}
  
  /*use full when you want to write sheet
   * where
   * You want to write in known sheet of the work book
   *  for this pre-requiste the work book,sheet should exists and
   *  atleast one row should be created
   *  
   *  Enhancement:
   *  If no sheet exists create a new sheet below method or separate method to create a sheet
   * */
  
  public void writeExcelData(ExcelReader excel1,String sheetName,int rowNo,int colNo,String dataToWrite){
	  Row row;
	  
	  try{
	  sheet=workbook.getSheet(sheetName);
	  if(excel1.rowCount(sheetName)>=rowNo){
	  row=sheet.getRow(rowNo-1);
	  }else{
		  row=sheet.createRow(rowNo-1);
	  } 
	  Cell cell=row.createCell(colNo);
	  cell.setCellValue(dataToWrite);
	  fout=new FileOutputStream(f);
      workbook.write(fout);
      fout.close();
	  }catch(IOException e){
	    	
	    	e.printStackTrace();
	    	System.out.println("Error accured while writing data in to the Cell");
	    }
  }
  
  /*
   * THis Method will help to create a new workbook,new sheet and write some thing in run time there
   */
  
  public void createNewSheetAndWrite(String sheetName,int rowNo,int colNo,String dataToWrite){
	  
	  Row row;
	  try{
	  XSSFWorkbook workbook=new XSSFWorkbook();
	  XSSFSheet sheet=workbook.createSheet(sheetName);
	  row=sheet.createRow(rowNo-1);
	  Cell cell=row.createCell(colNo);
	  cell.setCellValue(dataToWrite);
      fout=new FileOutputStream(f);
      workbook.write(fout);
      fout.close();  
	  }catch(IOException e){
	    	
	    	e.printStackTrace();
	    	System.out.println("Error accured while getting ColumnCount from the SheetName and RowNo provided");
	    }
  }
	
}
