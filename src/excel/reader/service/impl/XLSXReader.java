/**
 * 
 */
package excel.reader.service.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import excel.reader.IExcelReader;
import excel.util.ExcelUtils;


/**
 * @author LuanNgu
 *
 */
public class XLSXReader implements IExcelReader {

	private ExcelUtils _configs;
	private List<String[]> _content;
	private File _file;
	private String _joiner = ";";
	
	public XLSXReader()
	{
		_content = new ArrayList<>();
		_file = null;
		_configs = ExcelUtils.getInstance();
	}
	
	/**
	 * 
	 */
	public XLSXReader(String filename, ExcelUtils utils) {
		// TODO Auto-generated constructor stub
		_content = new ArrayList<>();
		_file = new File(filename);
		_configs = utils;
	}

	/* (non-Javadoc)
	 * @see upskills.fileimport.IExcelReader#fileReader(java.lang.String)
	 */
	@Override
	public List<String[]> fileReader(String filename) {
		// TODO Auto-generated method stub
		
		Workbook workbook = null;
		try {

			FileInputStream excelFile = new FileInputStream(_file);
			workbook = new XSSFWorkbook(excelFile);

			// Get data sheet by sheet name or index of sheet.
			Sheet datatypeSheet = workbook.getSheetAt(_configs.get_sheet_id());
			_content = readSheetXLSX(datatypeSheet);

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (workbook != null) {
				try {
					workbook.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}
		return _content;
	}

	private List<String[]> readSheetXLSX(Sheet datatypeSheet) {
		// TODO Auto-generated method stub
		Iterator<Row> iterator = datatypeSheet.iterator();
		List<String[]> lines = new ArrayList<>();
		while (iterator.hasNext()) {

			Row currentRow = iterator.next();
			Iterator<Cell> cellIterator = currentRow.iterator();
			StringBuilder str = new StringBuilder();
			while (cellIterator.hasNext()) {

				Cell currentCell = cellIterator.next();
				if (currentCell.getCellTypeEnum() == CellType.STRING) {
					str.append(currentCell.getStringCellValue()).append(_joiner);
				} else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
					str.append(currentCell.getNumericCellValue()).append(_joiner);
				} else if (currentCell.getCellTypeEnum() == CellType.FORMULA) {
					if (currentCell.getCachedFormulaResultTypeEnum() == CellType.STRING) {
						str.append(currentCell.getStringCellValue()).append(_joiner);
					} else if (currentCell.getCachedFormulaResultTypeEnum() == CellType.NUMERIC) {
						str.append(currentCell.getNumericCellValue()).append(_joiner);
					}
				}

			}
			lines.add(str.toString().split(_joiner));
		}
		return lines;
	}


}
