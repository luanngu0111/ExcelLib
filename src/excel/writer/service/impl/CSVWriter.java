/**
 * 
 */
package excel.writer.service.impl;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import excel.util.ExcelUtils;
import excel.writer.IExcelWriter;

/**
 * @author LuanNgu
 *
 */
public class CSVWriter implements IExcelWriter {

	private ExcelUtils _configs;
	private FileWriter _file;

	/**
	 * 
	 */
	public CSVWriter() {
		// TODO Auto-generated constructor stub
		_file = null;
		_configs = ExcelUtils.getInstance();
	}

	/**
	 * @throws IOException
	 * 
	 */
	public CSVWriter(String filename, ExcelUtils utl) throws IOException {
		// TODO Auto-generated constructor stub
		_file = new FileWriter(filename);
		_configs = utl;
	}

	/*
	 * (non-Javadoc)
	 * 
	 * @see excel.writer.IExcelWriter#fileWriter(java.util.List,
	 * java.lang.String)
	 */
	@Override
	public void fileWriter(List<String[]> content, String filename) {
		// TODO Auto-generated method stub
		String csvSplit = _configs.get_splitter();

		try {
			if (_file == null)
				_file = new FileWriter(filename);
			StringBuilder sb = new StringBuilder();
			if (csvSplit.trim() != "" && csvSplit.trim() != null) {
				if (content != null && content.size() > 0) {
					for (String[] obj : content) {
						sb.append(String.join(csvSplit, obj));
						sb.append("\n");
					}
				}
			}
			_file.append(sb.toString());
			_file.flush();
			_file.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
