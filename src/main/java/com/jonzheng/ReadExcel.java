package com.jonzheng;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 读取Excel每一列数据
 * @author Jon
 *
 */
public class ReadExcel {  
	public static String File = "d:/test.xlsx";

    public void read() throws IOException  {  
        InputStream stream = new FileInputStream(File);
        Workbook wb = null;  
        String fileType = File.substring(File.lastIndexOf(".")+1);
        if (fileType.equals("xls")) {  
            wb = new HSSFWorkbook(stream);  
        }  
        else if (fileType.equals("xlsx")) {  
            wb = new XSSFWorkbook(stream);  
        }  
        else {  
            System.out.println("File Format Error");  
        }  
        Sheet sheet1 = wb.getSheetAt(0);
        String ret = "";
        for (Row row : sheet1) {  
            for (Cell cell : row) {
                switch (cell.getCellType()) {  
                case Cell.CELL_TYPE_BLANK:  
                    ret = "";  
                    break;  
                case Cell.CELL_TYPE_BOOLEAN:  
                    ret = String.valueOf(cell.getBooleanCellValue());  
                    break;  
                case Cell.CELL_TYPE_ERROR:  
                    ret = null;  
                    break;  
                case Cell.CELL_TYPE_FORMULA:  
                	System.out.println("CELL_TYPE_FORMULA:"+cell);
                    break;  
                case Cell.CELL_TYPE_NUMERIC:  
                    if (DateUtil.isCellDateFormatted(cell)) {   
                        Date theDate = cell.getDateCellValue();
                        System.out.println("CELL_TYPE_FORMULA:"+theDate);
                    } else {   
                        ret = NumberToTextConverter.toText(cell.getNumericCellValue());  
                    }  
                    break;  
                case Cell.CELL_TYPE_STRING:  
                    ret = cell.getRichStringCellValue().getString();
                    break;  
                default:  
                    ret = null;  
                }
                System.out.print(ret+"\t");
            }
            System.out.println("\n");
        }  
    }  
    
    public static void main(String[] args) throws IOException {
    	new ReadExcel().read();  
    }  
}  