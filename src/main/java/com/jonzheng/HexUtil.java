package com.jonzheng;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 从Excel读取两列，将之转化为16进制，拼接为长字符输出
 * 拼接后+","接下一个
 * 例：C1E,10F,1D6
 * @author Jon
 *
 */
public class HexUtil {
	public static final String FILE = "d:/shengj.xlsx";  //输入文件
	public static final int CLO1 = 2;  //第一列
	public static final int CLO2 = 3;  //第二列

    public String getHexString() throws IOException  {  
        InputStream stream = new FileInputStream(FILE);
        Workbook wb = null;  
        String fileType = FILE.substring(FILE.lastIndexOf(".")+1);
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
        int num1 = 0;
        int num2 = 0;
        String text = "";
        StringBuilder sb = new StringBuilder();
        for (Row row : sheet1) {
            if(row.getRowNum()<1) continue; //注意第一行被跳过
            switch (row.getCell(CLO1).getCellType()) {  
            case Cell.CELL_TYPE_NUMERIC:
            	num1 = (int)row.getCell(CLO1).getNumericCellValue();
                break;
            case Cell.CELL_TYPE_STRING:
            	num1 = Integer.parseInt(row.getCell(CLO1).toString());
                break;
            }
            switch (row.getCell(CLO2).getCellType()) {
            case Cell.CELL_TYPE_NUMERIC:
            	num2 = (int)row.getCell(CLO2).getNumericCellValue();
                break;
            case Cell.CELL_TYPE_STRING:
            	num2 = Integer.parseInt(row.getCell(CLO2).toString());
                break;
            }
            //System.out.print(num1+"\t"+Integer.toHexString(num1).toUpperCase()+"\t"+num2+"\t"+Integer.toHexString(num2).toUpperCase()+"\n");
            sb.append(Integer.toHexString(num1).toUpperCase()).append(Integer.toHexString(num2).toUpperCase()).append(",");
        }
        text = sb.deleteCharAt(sb.length()-1).toString();
        System.out.println(text);
        writeToTxet(text);
        wb.close();
        return text;
    }
    
    public void writeToTxet(String hex) throws IOException{
    	FileOutputStream output = new FileOutputStream("d:/hexString.txt");
    	byte [] buff =new byte[]{};
    	buff = hex.getBytes();
    	output.write(buff, 0, buff.length);
    	output.close();
    }
    
    public static void main(String[] args) throws IOException {
    	new HexUtil().getHexString();
    }  
}
