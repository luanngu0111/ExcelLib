/**
 * 
 */
package excel.reader.service.impl;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import excel.reader.IExcelReader;
import excel.util.ExcelUtils;
import jxl.read.biff.BiffException;

/**
 * @author LuanNgu
 *
 */
public class XLSReader implements IExcelReader {

	private ExcelUtils _configs;
	private File _file;
	private List<String[]> _content;
	
	/**
	 * 
	 */
	public XLSReader() {
		// TODO Auto-generated constructor stub
		this._file = null;
		this._content = new ArrayList<>();
		this._configs = ExcelUtils.getInstance();
	}

	public XLSReader(String filename, ExcelUtils utils) {
		// TODO Auto-generated constructor stub
		this._file = new File(filename);
		this._configs = utils;
		this._content = new ArrayList<>();
	}


	public File get_file() {
		return _file;
	}

	public void set_file(File _file) {
		this._file = _file;
	}

	public List<String[]> get_content() {
		return _content;
	}

	public void set_content(List<String[]> _content) {
		this._content = _content;
	}

	/*
	 * (non-Javadoc)
	 * 
	 * @see upskills.fileimport.IExcelReader#fileReader(java.lang.String)
	 */
	@Override
	public List<String[]> fileReader(String filename) {
		// TODO Auto-generated method stub
		if (_file == null)
			_file = new File(filename);
		jxl.Workbook workbook = null;
		List<String[]> lines = new ArrayList<>();
		try {
			workbook = jxl.Workbook.getWorkbook(_file);
			jxl.Sheet sheet = workbook.getSheet(_configs.get_sheet_id());
			lines = readSheetXLS(sheet);

		} catch (IOException e) {
			e.printStackTrace();
		} catch (BiffException e) {
			e.printStackTrace();
		} finally {
			if (workbook != null) {
				workbook.close();
			}
		}

		return lines;
	}

	/**
	 * Read data in a sheet of XLS file
	 * 
	 * @param sheet:
	 *            data type of sheet
	 * @return Array LIst of data in sheet
	 */
	private static List<String[]> readSheetXLS(jxl.Sheet sheet) {
		List<String[]> lines = new ArrayList<>();
		int row_num = sheet.getRows();
		int col_num = sheet.getColumns();

		for (int i = 0; i < row_num; i++) {
			String[] row = new String[col_num];
			List<jxl.Cell> rows = Arrays.asList(sheet.getRow(i));
			row = (String[]) rows.stream().map(jxl.Cell::getContents).collect(Collectors.toList()).toArray();
			lines.add(row);
		}

		return lines;
	}

	
}
