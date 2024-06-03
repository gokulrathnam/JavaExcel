package ReadExcel;

import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;

import jxl.Cell;
import jxl.Workbook;
import jxl.Sheet;

public class ExcelReader {
	public static void main(String[] args) {
		try {
			Workbook workbook = Workbook.getWorkbook(new File("C:/Users/XXXX/XXXX/XXXX-V2.xls"));
			Sheet sheet = workbook.getSheet(0);
			String sQuery = "insert into test values (?,?,?,?,?,?,?,?,?,?,?,?)";
			//Class.forName("oracle.jdbc.driver.OracleDriver");
			//Connection conn = DriverManager.getConnection("jdbc:oracle:thin:@10.54.6.51:2021:CBIUATDB", "GBM", "gbm");
			//PreparedStatement statement = conn.prepareStatement(sQuery);

			for (int j = 1; j < sheet.getRows(); j++) {

				Cell cel1 = sheet.getCell(0, j);
				Cell cel2 = sheet.getCell(1, j);
				Cell cel3 = sheet.getCell(2, j);
				Cell cel4 = sheet.getCell(3, j);
				Cell cel5 = sheet.getCell(4, j);
				Cell cel6 = sheet.getCell(5, j);
				Cell cel7 = sheet.getCell(6, j);
				Cell cel8 = sheet.getCell(7, j);
				Cell cel9 = sheet.getCell(8, j);
				Cell cel10 = sheet.getCell(9, j);
				Cell cell1 = sheet.getCell(10, j);
				Cell cell2 = sheet.getCell(11, j);
				System.out.println(cel1+"----->"+cel1.getContents());
				System.out.println(cel2+"----->"+cel2.getContents());
				System.out.println(cel3+"----->"+cel3.getContents());
				System.out.println(cel4+"----->"+cel4.getContents());
				System.out.println(cel5+"----->"+cel5.getContents());
				System.out.println(cel6+"----->"+cel6.getContents());
				System.out.println(cel7+"----->"+cel7.getContents());
				System.out.println(cel8+"----->"+cel8.getContents());
				System.out.println(cel9+"----->"+cel9.getContents());
				System.out.println(cel10+"----->"+cel10.getContents());
				System.out.println(cell1+"----->"+cell1.getContents());
				System.out.println(cell2+"----->"+cell1.getContents());
				

				/*statement.setString(1, cel1.getContents());
				statement.setString(2, cel2.getContents());
				statement.setString(3, cel3.getContents());
				statement.setString(4, cel4.getContents());
				statement.setString(5, cel5.getContents());
				statement.setString(6, cel6.getContents());
				statement.setString(7, cel7.getContents());
				statement.setString(8, cel8.getContents());
				statement.setString(9, cel9.getContents());
				statement.setString(10, cel10.getContents());
				statement.setString(11, cell1.getContents());
				statement.setString(12, cell2.getContents());

				statement.executeUpdate();*/
			}
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}