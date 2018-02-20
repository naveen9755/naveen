import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class ReadWriteXLImple {
	static Boolean check = false;
	public static void main(String args[]) throws IOException {
		try {

			ArrayList arr1 = new ArrayList();
			ArrayList arr2 = new ArrayList();
			ArrayList arr3 = new ArrayList();

			FileInputStream file1 = new FileInputStream(new File("/Users/naveensingh/Documents/Jan_Timesheet3.xlsx"));

			FileInputStream file2 = new FileInputStream(new File("/Users/naveensingh/Documents/Jan_Timesheet3.xlsx"));

			// Get the workbook instance for XLS file
			HSSFWorkbook workbook1 = new HSSFWorkbook(file1);

			// Get first sheet from the workbook
			HSSFSheet sheet1 = workbook1.getSheetAt(0);
			HSSFSheet sheet2 = workbook1.getSheetAt(1);

			// Compare sheets

			// Get iterator to all the rows in current sheet1
			Iterator<Row> rowIterator1 = sheet1.iterator();
			Iterator<Row> rowIterator2 = sheet2.iterator();

			while (rowIterator1.hasNext()) {
				Row row = rowIterator1.next();
				// For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {

					Cell cell = cellIterator.next();

					// This is for read only one column from excel
					if (cell.getColumnIndex() == 0) {
						// Check the cell type and format accordingly
						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_NUMERIC:
							System.out.print(cell.getNumericCellValue());
							arr1.add(cell.getNumericCellValue());
							break;
						case Cell.CELL_TYPE_STRING:
							arr1.add(cell.getStringCellValue());
							System.out.print(cell.getStringCellValue());
							break;
						case Cell.CELL_TYPE_BOOLEAN:
							arr1.add(cell.getBooleanCellValue());
							System.out.print(cell.getBooleanCellValue());
							break;
						}

					}

				}

				System.out.println(" ");
			}

			file1.close();

			System.out.println("-----------------------------------");
			// For retrive the second excel data
			while (rowIterator2.hasNext()) {
				Row row1 = rowIterator2.next();
				// For each row, iterate through all the columns
				Iterator<Cell> cellIterator1 = row1.cellIterator();

				while (cellIterator1.hasNext()) {

					Cell cell1 = cellIterator1.next();
					// Check the cell type and format accordingly

					// This is for read only one column from excel
					if (cell1.getColumnIndex() == 0) {
						switch (cell1.getCellType()) {
						case Cell.CELL_TYPE_NUMERIC:
							arr2.add(cell1.getNumericCellValue());
							System.out.print(cell1.getNumericCellValue());
							break;
						case Cell.CELL_TYPE_STRING:
							arr2.add(cell1.getStringCellValue());
							System.out.print(cell1.getStringCellValue());
							break;
						case Cell.CELL_TYPE_BOOLEAN:
							arr2.add(cell1.getBooleanCellValue());
							System.out.print(cell1.getBooleanCellValue());
							break;

						}

					}
					// this continue is for
					// continue;
				}

				System.out.println("");
			}

			System.out.println("book1.xls -- " + arr1.size());
			System.out.println("book1.xls -- " + arr2.size());

			// compare two arrays
			for (Object process : arr1) {
				if (!arr2.contains(process)) {
					arr3.add(process);
				}
			}
			System.out.println("arr3 list values - = - = + " + arr3);
			writeStudentsListToExcel(arr3);

			// closing the files
			file1.close();
			file2.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	// write into new file excel

	private static void writeStudentsListToExcel(ArrayList arr3) {

		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream("/Users/naveensingh/Documents/Jan_Timesheet3.xlsx");

			HSSFWorkbook workBook = new HSSFWorkbook();
			HSSFSheet spreadSheet = workBook.createSheet("Sheet3");
			HSSFRow row;
			HSSFCell cell;
			// System.out.println("array size is :: "+minusArray.size());
			int cellnumber = 0;
			for (int i1 = 0; i1 < arr3.size(); i1++) {
				row = spreadSheet.createRow(i1);
				cell = row.createCell(cellnumber);
				// System.out.print(cell.getCellStyle());
				cell.setCellValue(arr3.get(i1).toString().trim());
			}
			workBook.write(fos);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}

		catch (IOException e) {
			e.printStackTrace();
		}

	}

	// end -write into new file
}
