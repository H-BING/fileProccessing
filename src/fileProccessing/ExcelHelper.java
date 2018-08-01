package fileProccessing;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Set;
import java.util.TreeSet;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ExcelHelper {
	
	public static Set<String> set = new TreeSet<String>();
	
	public static String[] addr = {
			"大田", "建宁", "将乐", "明溪", "宁化", "清流", "三明", "泰宁",
			"永安", "尤溪"
	};
	
	public static char[] level = {
			'县', '市'
	};
	
	
	public static void dealExcel(String filename) throws Exception {
		File file = new File(filename);
		
		//create workbook
		Workbook workbook = Workbook.getWorkbook(file);
		//get first sheet
		Sheet sheet = workbook.getSheet(0);
		//get data
		for(int i = 1; i < sheet.getRows(); i++) {
			Cell cell = sheet.getCell(0, i);
			//前三个字
			String substr3 = cell.getContents().substring(0, 3);
			//前两个字
			String substr2 = cell.getContents().substring(0, 2);

			if(!isAddrRight(substr3, substr2)) {
				set.add(cell.getContents());
			}
		}
		
		printSet();
		//outputToText(Main.FILE_OUTPUT);
		
		//shut down
		workbook.close();
	}
	
	public static void outputToText(String filepath) throws Exception {
		File file = new File(filepath);
		OutputStream outputStream = new FileOutputStream(file); 
		
		StringBuilder builder = new StringBuilder();
		if(set.isEmpty()) {
			builder.append("Empty!");
		} else {
			for(String value : set) {
				builder.append(value).append("\r\n");
			}	
		}
		byte[] data = builder.toString().getBytes();
        outputStream.write(data);
        outputStream.close();
	}
	
	public static void printSet() {
		if(set.isEmpty()) {
			System.out.println("empty!");
			return;
		}
		for(String value : set) {
			System.out.println(value);
		}
	}
	
	public static boolean isAddrRight(String sub3, String sub2) {
		if(isExist(sub2)) {
			for(int i = 0; i <level.length; i++) {
				if(level[i] == sub3.charAt(2)) {
					return true;
				}
			}
		} else if(sub2.equals("沙县")) {
			return true;
		}
		return false;
	}
	
	public static boolean isExist(String sub2) {
		for(int i = 0; i < addr.length; i++) {
			if(addr[i].equals(sub2)) {
				return true;
			}
		}
		return false;
	}
	
	public static File saveExcel(String filename) {
		String rootPah = "first\\file\\";
		File dir = new File(rootPah);
		return new File(dir, filename);
	}
	

}
