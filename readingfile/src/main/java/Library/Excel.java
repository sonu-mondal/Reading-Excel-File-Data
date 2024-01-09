package Library;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

//import com.microsoft.schemas.office.visio.x2012.main.CellType;

public class Excel {

	String excelFilePath="D:\\New folder\\excel1.xlsx";
	XSSFWorkbook wb;
	XSSFSheet sheet;
	public Excel() {
		try {
			FileInputStream fis=new FileInputStream(new File(this.excelFilePath));
			System.out.println("File input stream created successfully");
			wb=new XSSFWorkbook(fis);
			sheet=wb.getSheetAt(0);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void readSheetData() {
		Iterator<Row> rows=sheet.iterator();
		while(rows.hasNext()) {
			Row currentRow=rows.next();
			Iterator<Cell> cells=currentRow.cellIterator();
			while(cells.hasNext()) {
				Cell currentCell=cells.next();
				CellType cellType=currentCell.getCellType();
				
				String value="";
				if(cellType==cellType.STRING) {
					value=currentCell.getStringCellValue();
				}
				else if(cellType==CellType.NUMERIC){
					value=""+currentCell.getNumericCellValue();
				}
				
				System.out.println("Value for cell: "+value);
			}
		}
	}
	public static void main(String[] args) {
		Excel xl=new Excel();
		xl.readSheetData();
		
		
	}

}
