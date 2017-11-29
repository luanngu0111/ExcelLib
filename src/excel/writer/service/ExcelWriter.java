/**
 * 
 */
package excel.writer.service;

import java.util.List;

import excel.util.ExcelUtils;
import excel.writer.IExcelWriter;
import excel.writer.service.impl.CSVWriter;
import excel.writer.service.impl.XLSWriter;
import excel.writer.service.impl.XLSXWriter;

/**
 * @author LuanNgu
 *
 */
public class ExcelWriter {

	/**
	 * 
	 */
	public ExcelWriter() {
		// TODO Auto-generated constructor stub
	}
	
	public static void writeData(List<String[]> content, String filename, ExcelUtils utl) throws Exception
	{
		String ext = filename.substring(filename.lastIndexOf(".") + 1, filename.length());
		IExcelWriter writer = null;
		switch (ext) {
		case "csv":
			writer = new CSVWriter(filename, utl);
			break;
		case "xls":
			writer = new XLSWriter(filename, utl);
			break;
		case "xlsx":
			writer = new XLSXWriter(filename, utl);
			break;
		default:
			System.out.println(String.format("The format %s currently are not supported", ext));
			break;
		}
		if (writer != null)
		{
			writer.fileWriter(content, filename);
		}
	}

}
