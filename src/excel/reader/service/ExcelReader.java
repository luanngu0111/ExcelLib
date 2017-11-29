/**
 * 
 */
package excel.reader.service;

import java.util.List;

import excel.reader.IExcelReader;
import excel.reader.service.impl.CSVReader;
import excel.reader.service.impl.XLSReader;
import excel.reader.service.impl.XLSXReader;
import excel.util.ExcelUtils;


/**
 * @author LuanNgu
 *
 */
public class ExcelReader {

	/**
	 * 
	 */
	public ExcelReader() {
		// TODO Auto-generated constructor stub
	}
	
	public static List<String[]> readData(String filename, ExcelUtils utl)
	{
		String ext = filename.substring(filename.lastIndexOf(".") + 1, filename.length());
		IExcelReader reader;
		switch (ext) {
		case "csv":
			reader = new CSVReader(filename, utl);
			return reader.fileReader(filename);
		case "xls":
			reader = new XLSReader(filename, utl);
			return reader.fileReader(filename);
		case "xlsx":
			reader = new XLSXReader(filename, utl);
			return reader.fileReader(filename);
		default:
			System.out.println(String.format("The format %s are not supported", ext));
			return null;
		}
	}

}
