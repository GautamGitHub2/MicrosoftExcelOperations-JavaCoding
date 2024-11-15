package exceloperations;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

public class WorkingWithDateCells {
    public static void main(String[] args) throws IOException {

        //Create blank workbook
        XSSFWorkbook workbook=new XSSFWorkbook();

        //Create blank sheet
        XSSFSheet sheet= workbook.createSheet("Date Formats");

        //Format-0: Date in number format
        XSSFCell cell=sheet.createRow(0).createCell(0);
        cell.setCellValue(new Date());

        XSSFCreationHelper creationHelper=workbook.getCreationHelper();

        //Format-1: dd-mm-yyyy
        CellStyle style1= workbook.createCellStyle();
        style1.setDataFormat(creationHelper.createDataFormat().getFormat("dd-mm-yyyy")); //Specify the date format
        XSSFCell cell1=sheet.createRow(1).createCell(0);
        cell1.setCellValue(new Date());
        cell1.setCellStyle(style1);

        //Format-2: mm-dd-yyyy
        CellStyle style2= workbook.createCellStyle();
        style2.setDataFormat(creationHelper.createDataFormat().getFormat("mm-dd-yyyy")); //Specify the date format
        XSSFCell cell2=sheet.createRow(2).createCell(0);
        cell2.setCellValue(new Date());
        cell2.setCellStyle(style2);

        //Format-3: yyyy-mm-dd
        CellStyle style3= workbook.createCellStyle();
        style3.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-mm-dd")); //Specify the date format
        XSSFCell cell3=sheet.createRow(3).createCell(0);
        cell3.setCellValue(new Date());
        cell3.setCellStyle(style3);

        //Format-4: yyyy-dd-mm
        CellStyle style4= workbook.createCellStyle();
        style4.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-dd-mm")); //Specify the date format
        XSSFCell cell4=sheet.createRow(4).createCell(0);
        cell4.setCellValue(new Date());
        cell4.setCellStyle(style4);

        //Format-5: dd-mm-yyyy hh:mm:ss
        CellStyle style5= workbook.createCellStyle();
        style5.setDataFormat(creationHelper.createDataFormat().getFormat("dd-mm-yyyy hh:mm:ss")); //Specify the date format
        XSSFCell cell5=sheet.createRow(5).createCell(0);
        cell5.setCellValue(new Date());
        cell5.setCellStyle(style5);

        //Format-6: hh:mm:ss
        CellStyle style6= workbook.createCellStyle();
        style6.setDataFormat(creationHelper.createDataFormat().getFormat("hh:mm:ss")); //Specify the date format
        XSSFCell cell6=sheet.createRow(6).createCell(0);
        cell6.setCellValue(new Date());
        cell6.setCellStyle(style6);

        FileOutputStream fos=new FileOutputStream(".\\DataFiles\\DateFormats.xlsx");

        workbook.write(fos);
        workbook.close();
        fos.close();

        System.out.println("Date Formats workbook / sheet has been completed..!!");
    }
}
