package EpinionTechnicalTest.TechnicalTest;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App {
	public static String SAMPLE_XLSX_FILE_PATH = "D:/Input data.xlsx";
	public static void main(String[] args)
			throws EncryptedDocumentException, InvalidFormatException, IOException, ParseException {
		long startTime = System.nanoTime();

		// Creating a Workbook from an Excel file (.xls or .xlsx)
		Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));

		// Getting the Sheet at index zero
		Sheet sheet = workbook.getSheetAt(0);
		// Create a DataFormatter to format and get each cell's value as String
		DataFormatter dataFormatter = new DataFormatter();
		System.out.println("handling excel file.....");
		// set up parse string to date
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("d-MMM-yy");
		
		Workbook workbookwrite = new XSSFWorkbook();
		CreationHelper createHelper = workbookwrite.getCreationHelper();
		Sheet sheetwrite = workbookwrite.createSheet("Output");
		Font headerFont = workbookwrite.createFont();
		headerFont.setBold(true);
		CellStyle headerCellStyle = workbookwrite.createCellStyle();
        headerCellStyle.setFont(headerFont);
        //create excel header
        Row headerRow = sheetwrite.createRow(0);
        Cell cell = headerRow.createCell(0);
        cell.setCellValue("DateTime");
        cell.setCellStyle(headerCellStyle);
        Cell cell1 = headerRow.createCell(1);
        cell1.setCellValue("WeekDay");
        cell1.setCellStyle(headerCellStyle);
        Cell cell2 = headerRow.createCell(2);
        cell2.setCellValue("FlightNo");
        cell2.setCellStyle(headerCellStyle);
        Cell cell3 = headerRow.createCell(3);
        cell3.setCellValue("Destination");
        cell3.setCellStyle(headerCellStyle);
        Cell cell4 = headerRow.createCell(4);
        cell4.setCellValue("No.Passengers");
        cell4.setCellStyle(headerCellStyle);
        //create datetime cell style
        CellStyle dateCellStyle = workbookwrite.createCellStyle();
        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yyyy h:mm"));
        //create number cell style
        CellStyle numberCellStyle = workbookwrite.createCellStyle();
        numberCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("%"));
        int rownum = 1;
		for (int i=1;i<sheet.getPhysicalNumberOfRows();i++) {
			//get data from excel row
			String flightno = dataFormatter.formatCellValue(sheet.getRow(i).getCell(0));
			String airline = dataFormatter.formatCellValue(sheet.getRow(i).getCell(1));
			String datefrom = dataFormatter.formatCellValue(sheet.getRow(i).getCell(2));
			String dateto = dataFormatter.formatCellValue(sheet.getRow(i).getCell(3));
			String time = dataFormatter.formatCellValue(sheet.getRow(i).getCell(4));
			String destination = dataFormatter.formatCellValue(sheet.getRow(i).getCell(5));
			String Doop = dataFormatter.formatCellValue(sheet.getRow(i).getCell(6));
			String nopassenger = dataFormatter.formatCellValue(sheet.getRow(i).getCell(7));
			//datetime to handle
			LocalDate localdatefrom = LocalDate.parse(datefrom, formatter);
			LocalDate localdateto = LocalDate.parse(dateto, formatter);
			LocalDate localdatepoint = localdatefrom;
			while(localdatepoint.isBefore(localdateto) || localdatepoint.isEqual(localdateto)) {
				if(compareDOOP(Doop, localdatepoint.getDayOfWeek().toString())>-1) {
					
					Row row1 = sheetwrite.createRow(rownum++);
					Cell dateOfBirthCell = row1.createCell(0);
					dateOfBirthCell.setCellValue(localdatepoint.getMonthValue() + "/" + localdatepoint.getDayOfMonth() +"/" 
							+localdatepoint.getYear() +" "+ new StringBuilder(time).insert(2,":"));
		            dateOfBirthCell.setCellStyle(dateCellStyle);
		            String week = localdatepoint.getDayOfWeek().toString();
		            row1.createCell(1).setCellValue(week.substring(0, 1).toUpperCase()+week.substring(1,week.length()).toLowerCase());
		            row1.createCell(2).setCellValue(airline+flightno);
		            row1.createCell(3).setCellValue(destination);
		            Cell noPassengerCell = row1.createCell(4);
		            noPassengerCell.setCellValue(Integer.parseInt(nopassenger));
		            //noPassengerCell.setCellStyle(numberCellStyle);
				}
				localdatepoint = localdatepoint.plusDays(1);
			}
			
		}
		sheetwrite.autoSizeColumn(0);
		sheetwrite.autoSizeColumn(1);
		sheetwrite.autoSizeColumn(2);
		sheetwrite.autoSizeColumn(3);
		sheetwrite.autoSizeColumn(4);
		FileOutputStream fileOut = new FileOutputStream("Output data.xlsx");
        workbookwrite.write(fileOut);
        fileOut.close();
        workbookwrite.close();
		workbook.close();
		long endTime = System.nanoTime();
		System.out.println("Handling Completed");
		System.out.println("Took " + (double) (endTime - startTime) / 1000000000 + " s");
	}
	
	public static void handle_data(Row row) throws IOException {
		
	}
	
	public static int compareDOOP(String DOOP, String dayOfWeek) {
		if(dayOfWeek.equals("MONDAY")) {
			return DOOP.indexOf("1");
		}
		if(dayOfWeek.equals("TUESDAY")) {
			return DOOP.indexOf("2");
		}
		if(dayOfWeek.equals("WEDNESDAY")) {
			return DOOP.indexOf("3");
		}
		if(dayOfWeek.equals("THURSDAY")) {
			return DOOP.indexOf("4");
		}
		if(dayOfWeek.equals("FRIDAY")) {
			return DOOP.indexOf("5");
		}
		if(dayOfWeek.equals("SATURDAY")) {
			return DOOP.indexOf("6");
		}
		if(dayOfWeek.equals("SUNDAY")) {
			return DOOP.indexOf("7");
		}
		return -1;
	}
}
