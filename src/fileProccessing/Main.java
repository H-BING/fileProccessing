package fileProccessing;

public class Main {
	
	public static String FILE_INPUT = "E:\\031502411\\eclipse-workspace\\fileProccessing\\files\\test.xls";
	public static String FILE_OUTPUT = "E:\\031502411\\eclipse-workspace\\fileProccessing\\files\\out.txt";
	
	public static void main(String[] args) throws Exception {
		//System.out.println("hello world");
		ExcelHelper.dealExcel(FILE_INPUT);
	}

}
