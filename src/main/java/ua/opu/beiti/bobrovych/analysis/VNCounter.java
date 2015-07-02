package ua.opu.beiti.bobrovych.analysis;

import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

public class VNCounter extends Counter {
	
	private int currentRow;

	public void countVn(List<Double> inputData, String sheetName, double r) {
		CellStyle style2 = getWorkbook().createCellStyle();
		style2.setDataFormat(getWorkbook().createDataFormat().getFormat("#,##0.00"));
		CellStyle style4 = getWorkbook().createCellStyle();
		style4.setDataFormat(getWorkbook().createDataFormat()
				.getFormat("#,##0.0000"));
		setCurrentSheet(sheetName);
		currentRow = 0;
		Row row = getCurrentSheet().createRow(currentRow);
		row.createCell(1).setCellValue("Xi");
		currentRow = 1;
		double elemSum = 0;
		for (double elem : inputData) {
			row = getCurrentSheet().createRow(currentRow++);
			row.createCell(1).setCellValue(elem);
			elemSum += elem;
		}

		row = getCurrentSheet().createRow(currentRow++);
		double elemAvg = elemSum / inputData.size();
		row.createCell(0).setCellValue("Среднее");
		Cell cell = row.createCell(1);
		cell.setCellStyle(style2);
		cell.setCellValue(elemAvg);
		row = getCurrentSheet().getRow(0);
		row.createCell(2).setCellValue("Xi-Xavg");
		double[] dif = new double[inputData.size()];
		double sumDifSqr = 0;
		for (int i = 0; i < inputData.size(); i++) {
			dif[i] = inputData.get(i) - elemAvg;
			row = getCurrentSheet().getRow(i + 1);
			cell = row.createCell(2);
			cell.setCellStyle(style2);
			cell.setCellValue(dif[i]);
			sumDifSqr += Math.pow(dif[i], 2);
		}
		double[] cumulativeDif = new double[inputData.size()];
		double max = cumulativeDif[0];
		double min = cumulativeDif[0];
		for (int i = 0; i < cumulativeDif.length; i++) {
			if (i == 0) {
				cumulativeDif[i] = dif[i];
			} else {
				cumulativeDif[i] = cumulativeDif[i - 1] + dif[i];
			}
			if (cumulativeDif[i] > max) {
				max = cumulativeDif[i];
			}
			if (cumulativeDif[i] < min) {
				min = cumulativeDif[i];
			}
			row = getCurrentSheet().getRow(i + 1);
			cell = row.createCell(3);
			cell.setCellStyle(style2);
			cell.setCellValue(cumulativeDif[i]);
		}
		row = getCurrentSheet().getRow(currentRow - 1);
		row.createCell(2).setCellValue(" max = ");
		cell = row.createCell(3);
		cell.setCellStyle(style2);
		cell.setCellValue(max);
		row = getCurrentSheet().createRow(currentRow++);
		row.createCell(2).setCellValue(" min = ");
		cell = row.createCell(3);
		cell.setCellStyle(style2);
		cell.setCellValue(min);
		row.createCell(0).setCellValue("Станд. откл.");
		getCurrentSheet().autoSizeColumn(0);
		cell = row.createCell(1);
		cell.setCellStyle(style2);
		cell.setCellValue(Math.sqrt(sumDifSqr / (inputData.size() - 1)));
		row = getCurrentSheet().createRow(currentRow);
		row.createCell(2).setCellValue("R = ");
		double rMaxMin = max - min;
		cell = row.createCell(3);
		cell.setCellStyle(style2);
		cell.setCellValue(rMaxMin);
		row = getCurrentSheet().getRow(0);
		getCurrentSheet().autoSizeColumn(3);
		row.createCell(3).setCellValue("Накопленная");
		getCurrentSheet().autoSizeColumn(3);
		row = getCurrentSheet().getRow(0);
		row.createCell(8).setCellValue(" n =");
		row.createCell(9).setCellValue(inputData.size());
		row = getCurrentSheet().getRow(1);
		row.createCell(8).setCellValue(" r =");
		row.createCell(9).setCellValue(r);
		double x1 = 1.0 / 3.0;
		double x2 = 2.0 / 3.0;
		double k = Math.pow(1.5 * inputData.size(), x1)
				* Math.pow(2 * r / (1 - Math.pow(r, 2)), x2);
		row = getCurrentSheet().getRow(2);
		row.createCell(8).setCellValue(" k =");
		cell = row.createCell(9);
		cell.setCellStyle(style2);
		cell.setCellValue(k);
		int q = (int) k;
		row = getCurrentSheet().getRow(3);
		row.createCell(8).setCellValue(" q =");
		row.createCell(9).setCellValue(q);
		row = getCurrentSheet().getRow(0);
		row.createCell(4).setCellValue("Сумма произведений");
		getCurrentSheet().autoSizeColumn(4);
		double[] sum = new double[q];
		int top = 1;
		int bottom = inputData.size() - 1;
		for (int i = 0; i < q; i++) {
			int cur = top;
			for (int j = 0; j < bottom; j++) {
				sum[i] += dif[j] * dif[cur];
				cur++;
			}
			row = getCurrentSheet().getRow(i + 1);
			cell = row.createCell(4);
			cell.setCellStyle(style2);
			cell.setCellValue(sum[i]);
			top++;
			bottom--;
		}
		row = getCurrentSheet().getRow(0);
		row.createCell(5).setCellValue("j");
		row.createCell(6).setCellValue("w");
		row.createCell(7).setCellValue("w*СуммаПроизв");
		getCurrentSheet().autoSizeColumn(7);
		double[] w = new double[q];
		double[] wSum = new double[q];
		double t = 0;
		for (int j = 0; j < q; j++) {
			w[j] = 1.0 - ((j + 1.0) / (q + 1.0));
			wSum[j] = w[j] * sum[j];
			t += wSum[j];
			row = getCurrentSheet().getRow(j + 1);
			row.createCell(5).setCellValue(j + 1);
			cell = row.createCell(6);
			cell.setCellStyle(style4);
			cell.setCellValue(w[j]);
			cell = row.createCell(7);
			cell.setCellStyle(style2);
			cell.setCellValue(wSum[j]);
		}
		row = getCurrentSheet().getRow(q + 1);
		row.createCell(6).setCellValue(" T = ");
		cell = row.createCell(7);
		cell.setCellStyle(style2);
		cell.setCellValue(t);
		double sigmaSqr = (sumDifSqr + 2 * t) / inputData.size();
		row = getCurrentSheet().getRow(4);
		row.createCell(8).setCellValue(" sigma^2 = ");
		cell = row.createCell(9);
		cell.setCellStyle(style2);
		cell.setCellValue(sigmaSqr);
		getCurrentSheet().autoSizeColumn(9);
		double sigma = Math.sqrt(sigmaSqr);
		row = getCurrentSheet().getRow(5);
		row.createCell(8).setCellValue(" sigma = ");
		cell = row.createCell(9);
		cell.setCellStyle(style2);
		cell.setCellValue(sigma);
		double qN = rMaxMin / sigma;
		row = getCurrentSheet().getRow(6);
		row.createCell(8).setCellValue(" Qn = ");
		cell = row.createCell(9);
		cell.setCellStyle(style4);
		cell.setCellValue(qN);
		double vN = qN / Math.sqrt(inputData.size());
		row = getCurrentSheet().getRow(0);
		row.createCell(10).setCellValue(" Vn = ");
		cell = row.createCell(11);
		cell.setCellStyle(style4);
		cell.setCellValue(vN);
		incrementYearsCounter();
		addStatistics(sheetName, vN);
	}
}
