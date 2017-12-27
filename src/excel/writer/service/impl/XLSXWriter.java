/**
 * 
 */
package excel.writer.service.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import excel.util.ExcelFormat;
import excel.util.ExcelUtils;
import excel.writer.IExcelWriter;

/**
 * @author LuanNgu
 *
 */
public class XLSXWriter implements IExcelWriter {

	private ExcelUtils _configs;
	private File _file;
	private XSSFWorkbook _workbook;
	private XSSFSheet _sheet;

	/**
	 * 
	 */
	public XLSXWriter() {
		// TODO Auto-generated constructor stub
		_file = null;
		_configs = ExcelUtils.getInstance();
	}

	/**
	 * @throws IOException
	 * @throws InvalidFormatException
	 * 
	 */
	public XLSXWriter(String filename, ExcelUtils utl) {
		// TODO Auto-generated constructor stub
		_file = new File(filename);
		_configs = utl;

		if (utl.isOverwrite()) {
			_workbook = new XSSFWorkbook();
		} else {
			try {
				_workbook = new XSSFWorkbook(new FileInputStream(_file));
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				_workbook = new XSSFWorkbook();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

		_sheet = (_workbook.getSheet(_configs.get_sheet_name()) == null)
				? _workbook.createSheet(_configs.get_sheet_name()) : _workbook.getSheet(_configs.get_sheet_name());
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
		fileWriter(content, filename, _configs.get_format());
	}

	private void writeToSheet(List<String[]> content) {
		int rowNum = 0;
		// System.out.println("Creating excel");
//		Row headrow = _sheet.createRow(rowNum++);
		int colNum = 0;

		for (String[] str : content) {
			Row row = _sheet.createRow(rowNum++);
			colNum = 0;
			for (String field : str) {
				Cell cell = row.createCell(colNum++);
				cell.setCellValue(field);

			}
		}
	}

	/**
	 * @param content
	 * @param filename
	 * @param format
	 *            format of output excel
	 */
	public void fileWriter(List<String[]> content, String filename, ExcelFormat format) {
		writeToSheet(content);
		CellStyle cellStyle = _workbook.createCellStyle();
		XSSFFont font = _workbook.createFont();
		cellStyle.setFont(font);
		format.set_font_style(font);
		format.set_cell_style(cellStyle);
		format.set_sheet(_sheet);
		format.refresh();

		System.out.println("...Formating content");
		for (int i = 0; i <= _sheet.getLastRowNum(); i++) {
			Row row = _sheet.getRow(i);
			for (Cell cell : row) {
				cell.setCellStyle(format.get_cell_style());
			}
		}
		System.out.println("Finished formating");
		try {
			FileOutputStream outputStream = new FileOutputStream(filename);
			_workbook.write(outputStream);
			_workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
