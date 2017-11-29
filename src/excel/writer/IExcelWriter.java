package excel.writer;

import java.util.List;

import excel.util.ExcelUtils;

public interface IExcelWriter {
	public void fileWriter(List<String[]> content, String filename);
	
}
