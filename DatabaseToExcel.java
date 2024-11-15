package exceloperations;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;

public class DatabaseToExcel {

    public static void main(String[] args) throws SQLException, IOException {

        //Connect to the database
        Connection con=DriverManager.getConnection("jdbc:mysql://localhost:3306/gautam","GautamMySQL","Gautam@MySQL123");

        //Statement Query
        Statement stmt= con.createStatement();
        ResultSet rs=stmt.executeQuery("SELECT * FROM LOCATIONS");

        //Excel
        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet= workbook.createSheet("Locations Data");

        XSSFRow row= sheet.createRow(0);
        row.createCell(0).setCellValue("LOCATION_ID");
        row.createCell(1).setCellValue("STREET_ADDRESS");
        row.createCell(2).setCellValue("POSTAL_CODE");
        row.createCell(3).setCellValue("CITY");
        row.createCell(4).setCellValue("STATE_PROVINCE");
        row.createCell(5).setCellValue("COUNTRY_ID");

        int r=1;
        while (rs.next())
        {
            int locId=rs.getInt("LOCATION_ID");
            String streetAdd=rs.getString("STREET_ADDRESS");
            String postalCode=rs.getString("POSTAL_CODE");
            String city=rs.getString("CITY");
            String stateProvince=rs.getString("STATE_PROVINCE");
            String countryId=rs.getString("COUNTRY_ID");

            row= sheet.createRow(r++);

            row.createCell(0).setCellValue(locId);
            row.createCell(1).setCellValue(streetAdd);
            row.createCell(2).setCellValue(postalCode);
            row.createCell(3).setCellValue(city);
            row.createCell(4).setCellValue(stateProvince);
            row.createCell(5).setCellValue(countryId);

        }
        FileOutputStream fos=new FileOutputStream(".\\DataFiles\\Locations.xlsx");

        workbook.write(fos);
        workbook.close();
        fos.close();
        con.close();

        System.out.println("Done..!!");
    }
}
