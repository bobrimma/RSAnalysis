package ua.opu.beiti.bobrovych.analysis;

import java.io.FileNotFoundException;
import java.io.IOException;

public interface ExcelManager {
	
	public void createStatisticsSheet(String title);
	public void saveWorkbook(String url) throws FileNotFoundException, IOException;

}
