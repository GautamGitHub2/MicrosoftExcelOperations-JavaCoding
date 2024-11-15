package exceloperations;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class WritingExcelDemo1 {
    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet= workbook.createSheet("Emp Info");

        Object empdata[][]=
                {
                        {"EmpID","Name","Job","Contact No.","Address"},
                        {201,"Gautam Kumar","Test Automation Engineer",323456,"Kolkata"},
                        {202,"Gobinda Kumar","Shooter",3456777,"New Delhi"},
                        {203,"Gaurav Kumar","Shop Keeper",56788888,"Greater Noida"},
                        {204,"Nitoo Kumari","House Wife",76765544,"Jaipur"},
                        {205,"Purushottam Raj","Software Developer",887787667,"Gaya"},
                        {206,"Rohan Chadda","Test Engineer",78678999,"Delhi"},
                        {207,"John Martin","QA",323987456,"Pune"},

                };

        //Using For-Each Loop

        /*
        int rowCount=0;
        for (Object emp[]:empdata)
        {
            XSSFRow row= sheet.createRow(rowCount++);

            int columnCount=0;
            for (Object value:emp)
            {
                XSSFCell cell= row.createCell(columnCount++);
                if (value instanceof String)
                    cell.setCellValue((String) value);
                if (value instanceof Integer)
                    cell.setCellValue((Integer)value);
                if (value instanceof Boolean)
                    cell.setCellValue((Boolean) value);
            }

        }

         */

        //Using for loop

        int rows=empdata.length;
        int cols=empdata[0].length;

        System.out.println(rows);
        System.out.println(cols);

        for (int r=0;r<rows;r++)
        {
            //Create row in Excel sheet (Just before going to the cells /columns of the row)
            XSSFRow row= sheet.createRow(r);

            for (int c=0;c<cols;c++)
            {
                //Now before writing data into the cell, i have to create cells/columns for that row
                XSSFCell cell=row.createCell(c);

                //now cell is created, now i have to create data into that cell by taking data from the 2D Array that i have created above "Object empdata[][]"

                //Read those data and update in the cell

                Object value=empdata[r][c];

                //now take the data from 2D array and exactly copy to the excel sheet
                // before setting values to the excel sheet we have to check that the values are String, Interger or Boolean

                if (value instanceof String)
                    cell.setCellValue((String) value);
                if (value instanceof Integer)
                    cell.setCellValue((Integer) value);
                if (value instanceof Boolean)
                    cell.setCellValue((Boolean) value);
            }
        }

        String filePath=".\\DataFiles\\EmployeeData1.xlsx";
        FileOutputStream outputStream=new FileOutputStream(filePath);
        workbook.write(outputStream);
        outputStream.close();
        System.out.println("EmployeeData1.xlsx file has been written successfully...!!!");
    }
}
