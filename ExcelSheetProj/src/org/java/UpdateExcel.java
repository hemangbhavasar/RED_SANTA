package org.java;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UpdateExcel {
	
	
	int rownum = 0;

	/**
	 * @param args
	 */
	public void updateSheet(String fileName){
		
		try{
		
		FileInputStream file = new FileInputStream(new File(fileName));
		 
        //Create Workbook instance holding reference to .xlsx file
        XSSFWorkbook workbook = new XSSFWorkbook(file);
 
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
         
        //Get first/desired sheet from the workbook
        XSSFSheet sheet = workbook.getSheetAt(0);
 
        //Iterate through each rows one by one
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext())
        {
            Row row = rowIterator.next();
            rownum=row.getRowNum();
        }
//		}
//	    catch (Exception e)
//	    {
//	        e.printStackTrace();
//	    }
		
		//Blank workbook
		//XSSFWorkbook workbook = new XSSFWorkbook();
         
        //Create a blank sheet
		//XSSFSheet sheet = workbook.createSheet("Employee Data");
    	/*FirstName
    	LastName
    	DOB
    	Martial_status
    	Blood Groop
    	House Number
    	Address
    	Street Name
    	City
    	Neighborhood
    	StreetDirection
    	StreetSuffix
    	StreetType
    	ZipCode*/

        //This data needs to be written (Object[])
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[] {"ID", "FIRSTNAME", "LASTNAME" ,"DOB","MATERIAL STATUS" , "BLOOD GROOP" , "HOUSE NO." , "ADDRESS" , "STREET NAME" ,"CITY" , "NEIGHBORHOOD" ,"STREET DIRECTION" ,"STREET SUFFIX" , "STREET TYPE" ,"ZIP CODE"});
      //  data.put("2", new Object[] {1, "kamlesh", "chouhan"});
      //  data.put("3", new Object[] {2, "hemang", "bhavsar"});
     //   data.put("4", new Object[] {3, "rahul ", "tanna"});
        data.put("5", new Object[] {4, "Mastekeers", "Redsanta"});
          
        //Iterate over data and write to sheet
        Set<String> keyset = data.keySet();
        //int rownum = 0;
        for (String key : keyset)
        {
            Row row = sheet.createRow(++rownum);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
               Cell cell = row.createCell(cellnum++);
               if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
//        try
//        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File(fileName));
            workbook.write(out);
            out.close();
            System.out.println(fileName+ " written successfully on disk.");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
	}

}
