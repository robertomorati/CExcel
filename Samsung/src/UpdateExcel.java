/**
 * Create by: Roberto Guimarães Morati Junior
 * Java SE 8, 2015
 * Version 1
 * 
 * File that mekes changes inside a file excel.
 */

import java.io.*;
import java.math.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class UpdateExcel {

	private String sheet;
	private String dirFile;
	private String visto;
	private String serial;
	private int colVisto;
	private int colStatus;

	public UpdateExcel() {

	}

	public UpdateExcel(String sheet, String dirFile, String visto,
			String serial, int colVisto, int colStatus) {
		this.sheet = sheet;
		this.dirFile = dirFile;
		this.visto = visto;
		this.serial = serial;
		this.colVisto = colVisto;
		this.colStatus = colStatus;
	}

	public boolean validateSheet(String sheet, String dir) {
		FileInputStream fileInputStream;
		HSSFWorkbook workbook;

		try {
			fileInputStream = new FileInputStream(new File(dir));
			workbook = new HSSFWorkbook(fileInputStream);

			HSSFSheet ws = workbook.getSheet(sheet);

			fileInputStream.close();
			workbook.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			return false;
		} catch (IOException e) {
			// TODO Auto-generated catch block
			return false;
		} catch (NullPointerException e) {
			return false;
		}

		return true;
	}

	public boolean validateColVisto(String sheet, String dir, int colVisto) {
		FileInputStream fileInputStream;
		HSSFWorkbook workbook;

		try {
			fileInputStream = new FileInputStream(new File(dir));
			workbook = new HSSFWorkbook(fileInputStream);

			HSSFSheet ws = workbook.getSheet(sheet);
			int rowNum = ws.getLastRowNum() + 1;

			int colNum = ws.getRow(0).getLastCellNum();

			if (colVisto > colNum)
				return false;
			fileInputStream.close();
			workbook.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			return false;
		} catch (IOException e) {
			// TODO Auto-generated catch block
			return false;
		} catch (NullPointerException e) {
			return false;
		}

		return true;
	}

	public boolean validateColStatus(String sheet, String dir, int colStatus) {
		FileInputStream fileInputStream;
		HSSFWorkbook workbook;

		try {
			fileInputStream = new FileInputStream(new File(dir));
			workbook = new HSSFWorkbook(fileInputStream);

			HSSFSheet ws = workbook.getSheet(sheet);
			int rowNum = ws.getLastRowNum() + 1;

			int colNum = ws.getRow(0).getLastCellNum();

			fileInputStream.close();
			workbook.close();
			if (colStatus > colNum)
				return false;

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			return false;
		} catch (IOException e) {
			// TODO Auto-generated catch block
			return false;
		} catch (NullPointerException e) {
			return false;
		}

		return true;
	}

	public boolean updateSheet() {
		try {

			// Open file
			File file = new File(this.dirFile);
			FileInputStream fileInputStream = new FileInputStream(file);

			HSSFWorkbook workbook;
			try {
				// get file to work
				workbook = new HSSFWorkbook(fileInputStream);

				// get an sheet with specific name
				HSSFSheet ws = workbook.getSheet(this.sheet);

				// get number of last row
				int rowNum = ws.getLastRowNum() + 1;

				// get number columns
				int colNum = ws.getRow(0).getLastCellNum();

				// read column
				for (int i = 0; i < rowNum; i++) {
					HSSFRow row = ws.getRow(i);

					HSSFCell cell = row.getCell(0);

					if (cellToString(cell).equals(this.serial)) {
						HSSFCellStyle style = workbook.createCellStyle();
						style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
						style.setBorderTop(HSSFCellStyle.BORDER_THIN);
						style.setBorderRight(HSSFCellStyle.BORDER_THIN);
						style.setBorderLeft(HSSFCellStyle.BORDER_THIN);

						// update visto
						row.createCell(this.colVisto).setCellValue(this.visto);
						row.getCell(this.colVisto).setCellStyle(style);
						style = null;

						// update status
						row.createCell(this.colStatus).setCellValue("True");
						row.getCell(this.colStatus).setCellStyle(style);
						style = null;
						
						return true;
					} else {
						/**
						 * HSSFCellStyle style = workbook.createCellStyle();
						 * style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
						 * style.setBorderTop(HSSFCellStyle.BORDER_THIN);
						 * style.setBorderRight(HSSFCellStyle.BORDER_THIN);
						 * style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
						 * 
						 * row.createCell(this.colVisto).setCellValue(
						 * "Não encontrado!");
						 * row.getCell(this.colVisto).setCellStyle(style); style
						 * = null;
						 * 
						 * row.createCell(this.colStatus).setCellValue("False");
						 * row.getCell(this.colStatus).setCellStyle(style);
						 * style = null;
						 **/
					}

				}

				fileInputStream.close();

				FileOutputStream outputStream = new FileOutputStream(file);
				workbook.write(outputStream);

				outputStream.close();
				workbook.close();
				return false;
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (NullPointerException e) {
			e.printStackTrace();
		}
		return false;

	}

	public static String cellToString(HSSFCell cell) {

		int type;
		Object result = "";
		type = cell.getCellType();

		switch (type) {

		case 0:// numeric value in excel
			result = new BigDecimal(cell.getNumericCellValue());
			break;
		case 1: // string value in excel
			result = cell.getStringCellValue();
			break;
		case 2: // boolean value in excel
			result = cell.getBooleanCellValue();
			break;
		default:
			// throw new
			// RunTimeException("There are not support for this type of cell");
		}

		return result.toString();
	}
}