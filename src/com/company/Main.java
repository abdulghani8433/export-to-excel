package com.company;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

public class Main {

    public static void main(String[] args)throws Exception {

        // Step 2 : Load and Register drivers

        // Loading drivers using forName() method
        String jdbcURL="jdbc:mysql://localhost:3306/college";
        String username="root";
        String password="root";

        // Registering drivers using Driver Manager
        // Step 3: Establish. a connection
        Connection connection = DriverManager.getConnection(jdbcURL,username,password5);

        // Step 4: Proces the statement
        // Getting data from the table details
        Statement statement = connection.createStatement();
        ResultSet resultSet = statement.executeQuery(
                "select * from college");

        // Step 5: Execute a query
        // Create a workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        // Create a spreadsheet inside a workbook
        XSSFSheet spreadsheet1
                = workbook.createSheet("college db");
        XSSFRow row = spreadsheet1.createRow(1);
        XSSFCell cell;

        // Step 6: Process the results
        cell = row.createCell(1);
        cell.setCellValue("sid");

        cell = row.createCell(2);
        cell.setCellValue("sname");

        cell = row.createCell(3);
        cell.setCellValue("course");

        cell = row.createCell(4);
        cell.setCellValue("address");

        cell = row.createCell(5);
        cell.setCellValue("fee");


        // i=2 as we will start writing from the
        // 2'nd row
        int i = 2;

        while (resultSet.next()) {
            row = spreadsheet1.createRow(i);
            cell = row.createCell(1);
            cell.setCellValue(resultSet.getInt("sid"));

            cell = row.createCell(2);
            cell.setCellValue(resultSet.getString("sname"));

            cell = row.createCell(3);
            cell.setCellValue(resultSet.getString("course"));

            cell = row.createCell(4);
            cell.setCellValue(resultSet.getString("address"));

            cell = row.createCell(5);
            cell.setCellValue(resultSet.getString("fee"));

            i++;
        }

        // Local directory on computer
        FileOutputStream output = new FileOutputStream(new File(
                "/Users/abdul.ghani/Desktop/excel/student_database.xlsx"));

        // write
        workbook.write(output);

        // Step 7: Close the connection
        output.close();

        // Display message for successful compilation of
        // program
        System.out.println(
                "exceldatabase.xlsx written successfully");
    }
}

