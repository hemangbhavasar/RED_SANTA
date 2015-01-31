package org.java;




public class WriteExcel {
	
	public static String FILENAME= "D:\\Matekeer Donor.xlsx";

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		UpdateExcel updateExcel = new UpdateExcel();
		updateExcel.updateSheet(FILENAME);
	}

}
