package ua.opu.beiti.bobrovych.analysis;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

public class Counter implements ExcelManager {
	private HSSFWorkbook workbook;
	private HSSFSheet statSheet;
	private HSSFSheet currentSheet;
	private int yearsCounter;

	public Counter() {

		workbook = new HSSFWorkbook();
	}

	public HSSFWorkbook getWorkbook() {
		return workbook;
	}

	public HSSFSheet getStatSheet() {
		return statSheet;
	}

	public HSSFSheet getCurrentSheet() {
		return currentSheet;
	}

	public void setCurrentSheet(String sheetName) {
		currentSheet = getWorkbook().createSheet(sheetName);
	}

	public int getYearsCounter() {
		return yearsCounter;
	}

	protected void incrementYearsCounter() {
		yearsCounter++;
	}

	public void createStatisticsSheet(String title) {
		statSheet = getWorkbook().createSheet("Statistics");
		Row row = statSheet.createRow(0);
		row.createCell(1).setCellValue(title);
		row = statSheet.createRow(1);
		row.createCell(0).setCellValue("Year");
		row = statSheet.createRow(2);
		row.createCell(0).setCellValue("Value");
	}

	protected void addStatistics(String year, double value) {
		CellStyle style = getWorkbook().createCellStyle();
		style.setDataFormat(getWorkbook().createDataFormat().getFormat(
				"#,##0.000"));
		Row row = getStatSheet().getRow(1);
		row.createCell(yearsCounter).setCellValue(year);
		row = getStatSheet().getRow(2);
		Cell cell = row.createCell(yearsCounter);
		cell.setCellStyle(style);
		cell.setCellValue(value);
	}

	public void saveWorkbook(String url) {
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(new File(url));
			workbook.write(out);
			System.out.println("Excel " + url + " written successfully..");
		} catch (FileNotFoundException e) {
			JOptionPane.showMessageDialog(new JFrame(), url + "\n"
					+ "Файл занят другим процессом.",
					"Ошибка записи данных", JOptionPane.ERROR_MESSAGE);
		} catch (IOException e) {
			JOptionPane.showMessageDialog(new JFrame(), e.getMessage(),
					"Ошибка", JOptionPane.ERROR_MESSAGE);
		} finally {
			try {
				out.close();
			} catch (IOException e) {
				JOptionPane.showMessageDialog(new JFrame(), e.getMessage(),
						"Ошибка", JOptionPane.ERROR_MESSAGE);
			}
		}
	}
}
