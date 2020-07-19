package service;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import model.ProGrad;

//			Progression -1 
//Go to src/service. Open the ExcelGenerator and fill the logic inside the excelGenerate method.
//
//Stick to the instructions clearly. If you face any issue contact your mentor to get the guidance. 

public class ExcelGenerator {
	
	
	FileOutputStream out;
	
	public HSSFWorkbook excelGenerate(ProGrad prograd, List<ProGrad> list ) throws IOException {
		try {
			
			
			HSSFWorkbook hwb = new HSSFWorkbook();
			HSSFSheet sheet = hwb.createSheet("Sample sheet");
			//Create a new row in current sheet
			int rownum = -1,cellnum=-1;
			 String filename="Sample sheet";
			//Create a new cell in current row
			 Row  row1 = sheet.createRow(++rownum);
			    Cell celli = row1.createCell(++cellnum);
				//Set value to new value
			    	celli.setCellValue(prograd.getName());
			    	//create new cell
			        Cell celli1 = row1.createCell(++cellnum);
			      	celli1.setCellValue(prograd.getId());
			      	Cell celli2 = row1.createCell(++cellnum);
			      	celli2.setCellValue(prograd.getRate());
			    	Cell celli3 = row1.createCell(++cellnum);
			      	celli3.setCellValue(prograd.getComment());
			    	Cell celli4 = row1.createCell(++cellnum);
			      	celli4.setCellValue(prograd.getRecommend());

		
			
			for (ProGrad key: list) {
				 Row  row = sheet.createRow(++rownum);
			    Cell cell = row.createCell(++cellnum);
				//Set value to new value
			    	cell.setCellValue(key.getName());
			    	//create new cell
			        Cell cell1 = row.createCell(++cellnum);
			      	cell1.setCellValue(key.getId());
			      	Cell cell2 = row.createCell(++cellnum);
			      	cell2.setCellValue(key.getRate());
			    	Cell cell3 = row.createCell(++cellnum);
			      	cell3.setCellValue(key.getComment());
			    	Cell cell4 = row.createCell(++cellnum);
			      	cell4.setCellValue(key.getRecommend());
			      	cellnum=0;
		
			    }
			
			
			
			// Do not modify the lines given below
		
			out = new FileOutputStream(filename);
			hwb.write(out);
		
			return hwb;
			        }
	
		catch (Exception e) {
				e.printStackTrace();
			}
		finally {
			out.close();
		}
		return null;
	
		}
}


