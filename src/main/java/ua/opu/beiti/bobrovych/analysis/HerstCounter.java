package ua.opu.beiti.bobrovych.analysis;

import java.util.Set;
import java.util.SortedMap;
import java.util.TreeMap;



import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;


public class HerstCounter extends Counter {
	private int groups;
	private int elements;
	private double[][] data;
	private double[] avgByGroup;
	private double[] some;
	private double[] standartDeviation;
	private double[][] dynamicSeries;
	private double[] max;
	private double[] min;
	private double[] range;
	private double[] rS;
	private int currentRow;

	private final SortedMap<Integer, Double> rSGrForRegression = new TreeMap<Integer, Double>();

	public void countHerst(double[] inputData, String sheetName) {
		setCurrentSheet(sheetName);
		currentRow = 10;
		Row row = getCurrentSheet().createRow(++currentRow);
		row.createCell(0).setCellValue("Расчеты:");
		divideIntoGroups(32, inputData);
		divideIntoGroups(16, inputData);
		divideIntoGroups(8, inputData);
		divideIntoGroups(4, inputData);
		divideIntoGroups(2, inputData);
		int v = rSGrForRegression.size();
		double herst = (((double) v - 1) / v) * countRegression();
		incrementYearsCounter();
		addStatistics(sheetName, herst);
	}

	private void divideIntoGroups(int n, double[] inputData) {
		groups = n;
		elements = 256 / n;

		data = new double[elements][groups];
		avgByGroup = new double[groups];
		some = new double[groups];
		standartDeviation = new double[groups];
		dynamicSeries = new double[elements][groups];
		max = new double[groups];
		min = new double[groups];
		range = new double[groups];
		rS = new double[groups];

		for (int j = 0; j < groups; j++) {
			double sum = 0.0;
			for (int i = 0; i < elements; i++) {
				data[i][j] = inputData[(j + i) + j * (elements - 1)];
				sum += data[i][j];
			}
			avgByGroup[j] = sum / elements;
		}

		for (int j = 0; j < groups; j++) {
			for (int i = 0; i < elements; i++) {
				some[j] += Math.pow(data[i][j] - avgByGroup[j], 2);
				standartDeviation[j] = Math.sqrt(some[j] / (elements - 1));
			}
		}

		for (int j = 0; j < groups; j++) {
			for (int i = 0; i < elements; i++) {
				if (0 == i) {
					dynamicSeries[i][j] = inputData[(j + i) + j
							* (elements - 1)]
							- avgByGroup[j];
				} else {
					dynamicSeries[i][j] = dynamicSeries[i - 1][j]
							+ inputData[(j + i) + j * (elements - 1)]
							- avgByGroup[j];
				}
			}
		}

		Row row = getCurrentSheet().createRow(++currentRow);
		row.createCell(1).setCellValue(
				"Расчет для " + n + " групп по " + elements + " элементов");
		row = getCurrentSheet().createRow(++currentRow);
		CellStyle style = getWorkbook().createCellStyle();
		style.setDataFormat(getWorkbook().createDataFormat().getFormat(
				"#,##0.000"));

		Cell cell = row.createCell(0);

		cell.setCellValue("k");
		for (int j = 1; j < groups + 1; j++) {
			cell = row.createCell(j);
			cell.setCellValue("a" + j);
		}
		for (int i = 1; i < elements + 1; i++) {
			row = getCurrentSheet().createRow(currentRow + i);
			cell = row.createCell(0);
			cell.setCellValue(i);
			for (int j = 0; j < groups; j++) {
				cell = row.createCell(j + 1);
				cell.setCellStyle(style);
				cell.setCellValue(data[i - 1][j]);
			}
		}
		currentRow += (elements + 1);
		row = getCurrentSheet().createRow(currentRow);
		row.createCell(0).setCellValue("Среднее");
		for (int j = 0; j < groups; j++) {
			cell = row.createCell(j + 1);
			cell.setCellStyle(style);
			cell.setCellValue(avgByGroup[j]);
		}

		row = getCurrentSheet().createRow(++currentRow);
		row.createCell(0).setCellValue("Стандартное отклонение");
		getCurrentSheet().autoSizeColumn(0);
		for (int i = 0; i < elements; i++) {
			for (int j = 0; j < groups; j++) {
				cell = row.createCell(j + 1);
				cell.setCellStyle(style);
				cell.setCellValue(standartDeviation[j]);
			}
		}
		currentRow += 2;
		row = getCurrentSheet().createRow(currentRow);
		row.createCell(1).setCellValue("Временной ряд накопленныx отклонений");
		row = getCurrentSheet().createRow(++currentRow);
		row.createCell(0).setCellValue("k");
		for (int i = 1; i < groups + 1; i++) {
			cell = row.createCell(i);
			cell.setCellValue("a" + i);
		}
		for (int i = 1; i < elements + 1; i++) {
			row = getCurrentSheet().createRow(i + currentRow);
			cell = row.createCell(0);
			cell.setCellValue(i);
			for (int j = 0; j < groups; j++) {
				cell = row.createCell(j + 1);
				cell.setCellStyle(style);
				cell.setCellValue(dynamicSeries[i - 1][j]);
			}
		}

		for (int j = 0; j < groups; j++) {
			min[j] = dynamicSeries[0][j];
			max[j] = dynamicSeries[0][j];
			for (int i = 0; i < elements; i++) {
				if (dynamicSeries[i][j] > max[j]) {
					max[j] = dynamicSeries[i][j];
				}
				if (dynamicSeries[i][j] < min[j]) {
					min[j] = dynamicSeries[i][j];
				}
			}
		}
		currentRow += (elements + 1);
		row = getCurrentSheet().createRow(currentRow);
		row.createCell(0).setCellValue("max");
		for (int j = 0; j < groups; j++) {
			cell = row.createCell(j + 1);
			cell.setCellStyle(style);
			cell.setCellValue(max[j]);
		}
		row = getCurrentSheet().createRow(++currentRow);
		row.createCell(0).setCellValue("min");
		for (int j = 0; j < groups; j++) {
			cell = row.createCell(j + 1);
			cell.setCellStyle(style);
			cell.setCellValue(min[j]);
		}
		row = getCurrentSheet().createRow(++currentRow);
		row.createCell(0).setCellValue("Диапазон");
		for (int j = 0; j < groups; j++) {
			cell = row.createCell(j + 1);
			cell.setCellStyle(style);
			range[j] = max[j] - min[j];
			cell.setCellValue(range[j]);

		}
		double sumRS = 0.0;
		double rSGroup = 0.0;
		row = getCurrentSheet().createRow(++currentRow);
		row.createCell(0).setCellValue("R/S");
		for (int j = 0; j < groups; j++) {
			cell = row.createCell(j + 1);
			cell.setCellStyle(style);
			rS[j] = range[j] / standartDeviation[j];
			sumRS += rS[j];
			cell.setCellValue(rS[j]);
		}
		cell = row.createCell(groups + 1);
		cell.setCellStyle(style);
		cell.setCellValue(sumRS);
		row = getCurrentSheet().createRow(++currentRow);
		row.createCell(0).setCellValue("Для n = " + elements);
		row.createCell(1).setCellValue("R/S=");
		rSGroup = 1.0 / groups * sumRS;
		cell = row.createCell(2);
		cell.setCellStyle(style);
		cell.setCellValue(rSGroup);
		rSGrForRegression.put(elements, rSGroup);
		currentRow++;
	}

	public double countRegression() {
		currentRow = 0;
		Row row = getCurrentSheet().createRow(currentRow);
		row.createCell(0).setCellValue("Регрессия");
		row = getCurrentSheet().createRow(++currentRow);
		CellStyle style = getWorkbook().createCellStyle();
		style.setDataFormat(getWorkbook().createDataFormat().getFormat(
				"#,##0.000"));
		int celNum = 0;
		row.createCell(celNum).setCellValue("n");
		row.createCell(++celNum).setCellValue("R/S");
		row.createCell(++celNum).setCellValue("x");
		row.createCell(++celNum).setCellValue("y");
		row.createCell(++celNum).setCellValue("x*y");
		row.createCell(++celNum).setCellValue("x^2");
		int v = rSGrForRegression.size();
		row.createCell(++celNum).setCellValue("v = " + v);
		Set<Integer> keySet = rSGrForRegression.keySet();
		celNum = 0;
		double x, y, xy, x2;
		double sumX = 0.0;
		double sumY = 0.0;
		double sumXY = 0.0;
		double sumX2 = 0.0;
		Double value;
		Cell cell;
		for (Integer key : keySet) {
			row = getCurrentSheet().createRow(++currentRow);
			value = rSGrForRegression.get(key);
			cell = row.createCell(celNum);
			cell.setCellValue(key);
			cell = row.createCell(++celNum);
			cell.setCellStyle(style);
			cell.setCellValue(value);
			cell = row.createCell(++celNum);
			cell.setCellStyle(style);
			cell.setCellValue(x = Math.log(key));
			sumX += x;
			cell = row.createCell(++celNum);
			cell.setCellStyle(style);
			cell.setCellValue(y = Math.log(value));
			sumY += y;
			cell = row.createCell(++celNum);
			cell.setCellStyle(style);
			cell.setCellValue(xy = x * y);
			sumXY += xy;
			cell = row.createCell(++celNum);
			cell.setCellStyle(style);
			cell.setCellValue(x2 = Math.pow(x, 2));
			sumX2 += x2;
			celNum = 0;
		}
		row = getCurrentSheet().createRow(++currentRow);
		cell = row.createCell(2);
		cell.setCellStyle(style);
		cell.setCellValue(sumX);
		cell = row.createCell(3);
		cell.setCellStyle(style);
		cell.setCellValue(sumY);
		cell = row.createCell(4);
		cell.setCellStyle(style);
		cell.setCellValue(sumXY);
		cell = row.createCell(5);
		cell.setCellStyle(style);
		cell.setCellValue(sumX2);
		currentRow += 2;
		row = getCurrentSheet().createRow(currentRow);
		double h = (v * sumXY - sumX * sumY) / (v * sumX2 - Math.pow(sumX, 2));
		cell = row.createCell(0);
		cell.setCellValue("H = ");
		cell = row.createCell(1);
		cell.setCellStyle(style);
		cell.setCellValue(h);
		return h;
	}
}
