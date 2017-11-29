/**
 * 
 */
package excel.writer.service.impl;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;

import excel.util.ExcelUtils;
import excel.writer.IExcelWriter;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/**
 * @author LuanNgu
 *
 */
public class XLSWriter implements IExcelWriter {
	
	private ExcelUtils _configs;
	private File _file;
	/**
	 * 
	 */
	public XLSWriter() {
		// TODO Auto-generated constructor stub
		_file = null;
		_configs = ExcelUtils.getInstance();
	}
	
	/**
	 * 
	 */
	public XLSWriter(String filename, ExcelUtils utl) {
		// TODO Auto-generated constructor stub
		_file = new File(filename);
		_configs = utl;
	}

	/* (non-Javadoc)
	 * @see excel.writer.IExcelWriter#fileWriter(java.util.List, java.lang.String)
	 */
	@Override
	public void fileWriter(List<String[]> content, String filename) {
		// TODO Auto-generated method stub
		String sheetname = _configs.get_sheet_name();
		int col_size = 0;
		// 1. Create an Excel file
		WritableWorkbook Wbook = null;
		try {
			if (_file==null)
				_file = new File(filename);
			Wbook = Workbook.createWorkbook(_file);
			// create an Excel sheet
			WritableSheet excelSheet = Wbook.createSheet(sheetname, 0);
			Label label = null;
			String[] str = null;
		
			for (int i = 0; i < content.size(); i++) {
				str = content.get(i);
				col_size = str.length;
				for (int j = 0; j < col_size; j++) {
					label = new Label(j, i, str[j]);
					excelSheet.addCell(label);
				}

			}
		} catch (IOException e) {
			e.printStackTrace();
		} catch (WriteException e) {
			e.printStackTrace();
		} finally {
			if (Wbook != null) {
				try {
					Wbook.close();
				} catch (IOException e) {
					e.printStackTrace();
				} catch (WriteException e) {
					e.printStackTrace();
				}
			}
		}
	}

}
