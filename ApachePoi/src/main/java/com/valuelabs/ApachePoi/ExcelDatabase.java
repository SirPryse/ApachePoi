package com.valuelabs.ApachePoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDatabase {
	public static void main(String[] args) throws Exception {
		Class.forName("com.mysql.jdbc.Driver");
		Connection connect = DriverManager.getConnection("jdbc:mysql://localhost:3306/Excel", "root", "");
		
		 
		FileInputStream fip = new FileInputStream(new File("EmployeeDB.xlsx"));
		  
		
		Statement statement = connect.createStatement();
		ResultSet resultSet = statement.executeQuery("select * from users");
		XSSFWorkbook workbook = new XSSFWorkbook(fip); 
		XSSFSheet spreadsheet = workbook.getSheetAt(0);

		XSSFRow row = spreadsheet.createRow(0);
		XSSFCell cell;
		cell = row.createCell(0);
		cell.setCellValue("EmpId");
		cell = row.createCell(1);
		cell.setCellValue("Name");
		cell = row.createCell(2);
		cell.setCellValue("Age");
		cell = row.createCell(3);
		cell.setCellValue("Country");
		cell = row.createCell(4);
		cell.setCellValue("State");
		cell = row.createCell(5);
		cell.setCellValue("City");
		int i = 1;

		while (resultSet.next()) {
			row = spreadsheet.createRow(i);
			cell = row.createCell(0);
			cell.setCellValue(resultSet.getInt("EmpId"));
			cell = row.createCell(1);
			cell.setCellValue(resultSet.getString("Name"));
			cell = row.createCell(2);
			cell.setCellValue(resultSet.getInt("Age"));
			cell = row.createCell(3);
			cell.setCellValue(resultSet.getString("Country"));
			cell = row.createCell(4);
			cell.setCellValue(resultSet.getString("State"));
			cell = row.createCell(5);
			cell.setCellValue(resultSet.getString("City"));
			i++;
		}
		resultSet.close();

		List<String> degree = new ArrayList<>();
		ResultSet uniquePost = statement.executeQuery("select DISTINCT EmpId from users");
		while (uniquePost.next()) {
			degree.add(uniquePost.getString("EmpId"));
		}
		uniquePost.close();
		String[] nameClass = new String[degree.size()];
		i = 0;
		for (String name : degree) {
			nameClass[i] = name;
			i++;
		}

		DataValidationHelper validationHelper = new XSSFDataValidationHelper(spreadsheet);
		CellRangeAddressList addressList = new CellRangeAddressList(1, 5, 0, 0);
		DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(nameClass);
		DataValidation dataValidation = validationHelper.createValidation(constraint, addressList);
		dataValidation.setSuppressDropDownArrow(true);
		spreadsheet.addValidationData(dataValidation);

		degree.clear();

		ResultSet uniquePost2 = statement.executeQuery("select DISTINCT Country from users");
		while (uniquePost2.next()) {
			degree.add(uniquePost2.getString("Country"));
		}
		uniquePost2.close();
		String[] nameClass2 = new String[degree.size()];
		i = 0;
		for (String name : degree) {
			nameClass2[i] = name;
			i++;
		}

		DataValidationHelper validationHelper2 = new XSSFDataValidationHelper(spreadsheet);
		CellRangeAddressList addressList2 = new CellRangeAddressList(1, 5, 3, 3);
		DataValidationConstraint constraint2 = validationHelper2.createExplicitListConstraint(nameClass2);
		DataValidation dataValidation2 = validationHelper2.createValidation(constraint2, addressList2);
		dataValidation2.setSuppressDropDownArrow(true);
		spreadsheet.addValidationData(dataValidation2);

		degree.clear();

		ResultSet uniquePost3 = statement.executeQuery("select DISTINCT State from users");
		while (uniquePost3.next()) {
			degree.add(uniquePost3.getString("State"));
		}
		uniquePost3.close();
		String[] nameClass3 = new String[degree.size()];
		i = 0;
		for (String name : degree) {
			nameClass3[i] = name;
			i++;
		}

		DataValidationHelper validationHelper3 = new XSSFDataValidationHelper(spreadsheet);
		CellRangeAddressList addressList3 = new CellRangeAddressList(1, 5, 4, 4);
		DataValidationConstraint constraint3 = validationHelper3.createExplicitListConstraint(nameClass3);
		DataValidation dataValidation3 = validationHelper3.createValidation(constraint3, addressList3);
		dataValidation3.setSuppressDropDownArrow(true);
		spreadsheet.addValidationData(dataValidation3);

		degree.clear();

		ResultSet uniquePost4 = statement.executeQuery("select DISTINCT City from users");
		while (uniquePost4.next()) {
			degree.add(uniquePost4.getString("City"));
		}
		uniquePost4.close();
		String[] nameClass4 = new String[degree.size()];
		i = 0;
		for (String name : degree) {
			nameClass4[i] = name;
			i++;
		}

		DataValidationHelper validationHelper4 = new XSSFDataValidationHelper(spreadsheet);
		CellRangeAddressList addressList4 = new CellRangeAddressList(1, 5, 5, 5);
		DataValidationConstraint constraint4 = validationHelper4.createExplicitListConstraint(nameClass4);
		DataValidation dataValidation4 = validationHelper4.createValidation(constraint4, addressList4);
		dataValidation4.setSuppressDropDownArrow(true);
		spreadsheet.addValidationData(dataValidation4);

		FileOutputStream out = new FileOutputStream(new File("EmployeeDB.xlsx"));
		workbook.write(out);
		workbook.close();
		out.close();
		System.out.println("EmployeeDB.xlsx written successfully");
	}
}