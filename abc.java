package com;

import java.io.File;
import java.io.FileInputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

	static XSSFWorkbook wb;
	
	public static void main(String[] args) {
	    try {
	        File file = new File("C:\\Users\\Rahul Khurana\\OneDrive\\Documents\\homer.xlsx");   //creating a new file instance
	        FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file
	        //creating Workbook instance that refers to .xlsx file
	        wb = new XSSFWorkbook(fis);
	        XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object
	        Row rw = sheet.getRow(0);
	        
	        int lastcell=sheet.getRow(0).getLastCellNum();
	        //Non empty Last cell Number or index return

	        LinkedHashMap<String,String> innerMap = new LinkedHashMap<String,String>();
	        LinkedHashMap<String,LinkedHashMap<String,String>> outerMap = new LinkedHashMap<String,LinkedHashMap<String,String>>();
	        
	        ArrayList<String> keysList =  new ArrayList<String>();
	        
	        for(int i=0;i<lastcell;i++)
	        {
	        	Cell cell = rw.getCell(i);
//	            innerMap.put(cell.getSheet().getRow(0).getCell(i).getRichStringCellValue().toString(), "");
	        	keysList.add(cell.getSheet().getRow(0).getCell(i).getRichStringCellValue().toString());
	        	
	        }
	        
	        Iterator<Row> iterator = sheet.iterator();

	        int colCounter;
            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                
                if(currentRow.getRowNum()==0){
                	continue;
                }
               
                colCounter=0;
                String paymentID="";
                
                for(String key : keysList) {
                	Cell cell = currentRow.getCell(colCounter);
                	innerMap.put(key, getCellValue(cell));
                	colCounter++;
                }
                paymentID = innerMap.get("roll");
                outerMap.put(paymentID, innerMap);
                innerMap = new LinkedHashMap<String,String>();
                
            }
            
	        fis.close();
	        
			System.out.println(outerMap);

	    } catch (Exception e) {
	        e.printStackTrace();
	    }
	}

	public static String getCellValue(Cell cell) {
		      String retVal;
		      if (cell == null) {
		          return "";
		      }
		      switch (cell.getCellType()) {
		          case Cell.CELL_TYPE_BLANK:
		              retVal = "";
		              break;
		 
		          case Cell.CELL_TYPE_BOOLEAN:
		              retVal = "" + cell.getBooleanCellValue();
		              break;
		
		          case Cell.CELL_TYPE_STRING:
		              retVal = cell.getStringCellValue();
		              break;
		
		          case Cell.CELL_TYPE_NUMERIC:
		              retVal = isNumberOrDate(cell);
		              break;
		
		          case Cell.CELL_TYPE_FORMULA:
		              retVal = processFormula(cell);
		              break;
		
		          default:
		              retVal = "";
		      }
		      return retVal;
		  }
	
	private static String processFormula(Cell cell) {
		        String retVal = "";
		        FormulaEvaluator evaluator =
		 wb.getCreationHelper().createFormulaEvaluator();
		        if (cell.getCachedFormulaResultType() == Cell.CELL_TYPE_ERROR) {
		            retVal = "#VALUE!";
		        } else {
		            retVal = formatter.formatCellValue(cell, evaluator);
		            if (retVal.matches("[0-9]+")) {
		                if (HSSFDateUtil.isCellDateFormatted(cell)) {
		                    retVal = dateFormat.format(cell.getDateCellValue());
		                } else {
		                    retVal = formatter.formatCellValue(cell);
		                }
		            }
		        }
		        return retVal;
		    }
	
	private static String isNumberOrDate(Cell cell) {
		        String retVal;
		        if (HSSFDateUtil.isCellDateFormatted(cell)) {
		            retVal = dateFormat.format(cell.getDateCellValue());
		 
		        } else {
		            retVal = formatter.formatCellValue(cell);
		        }
		        return retVal;
		    }
	private static DataFormatter formatter = new DataFormatter();
	private static DateFormat  dateFormat = new SimpleDateFormat("dd/MM/yyyy");
	
}
