package Library2;

import Library.Excel;

public class Usage {
	
	public static void main(String[] args) {
		Excel2 xl = new Excel2("D:\\New folder\\excel1.xlsx");
		xl.setSheet("Data");
		//xl.readSheetData();
		String lastName=xl.getCellData(6, 1);
		System.out.println("last name is : "+lastName);
	
	}
}
