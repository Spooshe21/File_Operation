/********************************************************************************************
 *   COPYRIGHT (C) 2024 CREVAVI TECHNOLOGIES PVT LTD
 *   The reproduction, transmission or use of this document/file or its
 *   contents is not permitted without written authorization.
 *   Offenders will be liable for damages. All rights reserved.
 *---------------------------------------------------------------------------
 *   Purpose:  Create an Excel file with employee data
 *   Project:  Excel Data Writing
 *   Platform: Cross-platform (Windows, macOS, Linux)
 *   Compiler: JDK-22
 *   IDE:      Eclipse IDE for Enterprise Java and Web Developers (includes Incubating components)
 *	           Version: 2024-03 (4.31.0)
 *             Build id: 20240307-1437
 ********************************************************************************************/

package file;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;

public class Operation {

	public static void main(String[] args) {
		// Attempt to create an Excel file and write data to it
		try {
			// Create a new workbook
			Workbook workbook = new XSSFWorkbook();

			// Create a new sheet named "Employee Data"
			Sheet sheet = workbook.createSheet("Employee Data");

			// Define the data to be written to the sheet
			Object[][] data = { { "ID", "NAME", "LASTNAME" }, { 1, "John", "Doe" }, { 2, "Anna", "Smith" },
					{ 3, "Peter", "Jones" } };

			// Write data to the sheet
			int rowNum = 0;
			for (Object[] aData : data) {
				Row row = sheet.createRow(rowNum++);
				int colNum = 0;
				for (Object obj : aData) {
					Cell cell = row.createCell(colNum++);
					if (obj instanceof String) {
						cell.setCellValue((String) obj);
					} else if (obj instanceof Integer) {
						cell.setCellValue((Integer) obj);
					}
				}
			}

			// Create a file output stream to write the workbook to a file
			FileOutputStream out = new FileOutputStream("employee.xlsx");
			workbook.write(out);
			out.close();
			workbook.close();

			// Print success message
			System.out.println("Excel file has been written successfully.");

		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
