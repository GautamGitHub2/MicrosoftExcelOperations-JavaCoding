package exceloperations;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class HashMapToExcel {
    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet= workbook.createSheet("Student Data");

        Map<String, String> data=new HashMap<String, String>();

        data.put("101","Gautam");
        data.put("102","Kumar");
        data.put("103","Purushottam");
        data.put("104","Raj");
        data.put("105","Nitoo");
        data.put("106","Kumari");
        data.put("107","Aman");
        data.put("108","Gupta");
        data.put("109","Gobinda");
        data.put("110","John");

        int rowno=0;

        for (Map.Entry entry: data.entrySet())
        {
            XSSFRow row= sheet.createRow(rowno++);

            row.createCell(0).setCellValue((String) entry.getKey());
            row.createCell(1).setCellValue((String) entry.getValue());
        }

        FileOutputStream fos=new FileOutputStream(".\\DataFiles\\StudentDetails.xlsx");
        workbook.write(fos);
        fos.close();
        System.out.println("Excel Sheet of Student Details has been written successfully..!!");

    }
}
