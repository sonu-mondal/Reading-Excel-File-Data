package Library2;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import Library.Excel;

public class Excel2 {
	String excelFilePath = "";
	XSSFWorkbook wb;
	XSSFSheet sheet;

	public Excel2(String excelFilePath) {

		try {

			this.excelFilePath = excelFilePath;
			FileInputStream fis = new FileInputStream(new File(this.excelFilePath));
			System.out.println("File input stream created successfully");
			wb = new XSSFWorkbook(fis);
//			sheet = wb.getSheetAt(0);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void setSheet(String sheetName) {
		sheet = wb.getSheet(sheetName);
	}

	public String getCellData(int rowNum, int colNum) {
		String ret = "";
		try {
			Row row = sheet.getRow(rowNum);
			Cell cell = row.getCell(colNum);

			if (cell.getCellType() == CellType.STRING) {
				ret = cell.getStringCellValue();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return ret;
	}

	public void readSheetData() {
		Iterator<Row> rows = sheet.iterator();
		while (rows.hasNext()) {
			Row currentRow = rows.next();
			Iterator<Cell> cells = currentRow.cellIterator();
			while (cells.hasNext()) {
				Cell currentCell = cells.next();
				CellType cellType = currentCell.getCellType();

				String value = "";
				if (cellType == cellType.STRING) {
					value = currentCell.getStringCellValue();
				} else if (cellType == CellType.NUMERIC) {
					value = "" + currentCell.getNumericCellValue();
				}

				System.out.println("Value for cell: " + value);
			}
		}
	}

}
