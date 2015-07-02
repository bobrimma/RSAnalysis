package ua.opu.beiti.bobrovych.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import ua.opu.beiti.bobrovych.exceptions.BlankCellException;
import ua.opu.beiti.bobrovych.exceptions.NotEnoughtDataException;
import ua.opu.beiti.bobrovych.exceptions.TooMuchDataException;

public class ExcelParser {
	private double[][] inputData;
	private ArrayList<String> years = new ArrayList<String>();
	private List<List<Double>> data = new ArrayList<List<Double>>();
	private ArrayList<Double> r = new ArrayList<Double>();

	public List<List<Double>> getData() {
		return data;
	}

	public List<Double> getR() {
		return r;
	}

	public void printYears() {
		System.out.println(years.toString());

	}

	public void printData() {
		for (int i = 0; i < years.size(); i++) {
			System.out.println(Arrays.toString(inputData[i]));
		}
	}

	public double[][] getInputData() {
		return inputData.clone();
	}

	@SuppressWarnings("unchecked")
	public ArrayList<String> getYears() {
		return (ArrayList<String>) years.clone();
	}

	public void parseExcelForHerst(String url) throws NotEnoughtDataException,
			TooMuchDataException, BlankCellException {
		FileInputStream file = null;
		try {
			file = new FileInputStream(new File(url));
			HSSFWorkbook workbook = new HSSFWorkbook(file);
			HSSFSheet sheet = workbook.getSheetAt(0);
			int rowNumbers = sheet.getLastRowNum() - sheet.getFirstRowNum();

			if (rowNumbers < 256) {
				throw new NotEnoughtDataException("Недостаточно данных.");

			} else if (rowNumbers > 256) {
				throw new TooMuchDataException("Слишком много данных.");

			} else {
				Iterator<Row> rowIterator = sheet.iterator();
				int i = 0;
				while (rowIterator.hasNext()) {
					Row row = rowIterator.next();

					Iterator<Cell> cellIterator = row.cellIterator();
					int j = 0;
					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						if (i == 0) {
							switch (cell.getCellType()) {
							case Cell.CELL_TYPE_NUMERIC:
								years.add(String.valueOf((int) cell
										.getNumericCellValue()));
								break;
							case Cell.CELL_TYPE_STRING:
								years.add(cell.getStringCellValue());
								break;
							}
							inputData = new double[years.size()][256];

						} else {
							if (cell.getCellType() != Cell.CELL_TYPE_BLANK) {
								inputData[j][i - 1] = cell
										.getNumericCellValue();
							} else {
								throw new BlankCellException("Ячейка (" + i
										+ "," + j + ") пустая.");
							}
						}
						j++;
					}
					i++;
				}
			}
		} catch (FileNotFoundException e) {
			JOptionPane.showMessageDialog(new JFrame(), url, "Файл не найден",
					JOptionPane.ERROR_MESSAGE);
		} catch (IOException e) {
			JOptionPane.showMessageDialog(new JFrame(), e.getMessage(),
					"Ошибка", JOptionPane.ERROR_MESSAGE);
		} finally {
			try {
				file.close();
			} catch (IOException e) {
				JOptionPane.showMessageDialog(new JFrame(), e.getMessage(),
						"Ошибка", JOptionPane.ERROR_MESSAGE);
			}
		}
	}

	public void parseExcelForVn(String url) throws NotEnoughtDataException {
		FileInputStream file = null;
		try {
			file = new FileInputStream(new File(url));
			HSSFWorkbook workbook = new HSSFWorkbook(file);
			HSSFSheet sheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = sheet.iterator();
			List<Double> currentList;
			int i = 0;
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				int j = 0;
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();

					if (i == 0) {
						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_NUMERIC: {
							years.add(String.valueOf((int) cell
									.getNumericCellValue()));
							data.add(new ArrayList<Double>());
							break;
						}
						case Cell.CELL_TYPE_STRING: {
							years.add(cell.getStringCellValue());
							data.add(new ArrayList<Double>());
							break;
						}
						}
					} else {
						if (cell.getCellType() != Cell.CELL_TYPE_BLANK) {
							currentList = data.get(j);
							currentList.add(cell.getNumericCellValue());
						}
						
					}
					j++;
				}
				i++;
			}
			sheet = workbook.getSheetAt(1);
			if (sheet != null) {
				Row row = sheet.getRow(1);
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					r.add(cell.getNumericCellValue());
				}
				if (r.size() != years.size()) {
					throw new NotEnoughtDataException(
							"Количество столбцов на первом и втором листе "
									+ "\n" + "должно совпадать.");
				}
			} else {
				throw new NotEnoughtDataException(
						"Отсутствует лист с коэффициентами автокорреляции.");
			}

		} catch (FileNotFoundException e) {
			JOptionPane.showMessageDialog(new JFrame(), url, "Файл не найден",
					JOptionPane.ERROR_MESSAGE);
		} catch (IOException e) {
			JOptionPane.showMessageDialog(new JFrame(), e.getMessage(),
					"Ошибка", JOptionPane.ERROR_MESSAGE);
		} finally {
			try {
				file.close();
			} catch (IOException e) {

				JOptionPane.showMessageDialog(new JFrame(), e.getMessage(),
						"Ошибка", JOptionPane.ERROR_MESSAGE);
			}
		}
	}

}
