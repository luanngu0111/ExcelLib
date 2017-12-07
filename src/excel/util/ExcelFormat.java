/**
 * 
 */
package excel.util;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;

/**
 * @author LuanNgu
 *
 */
public class ExcelFormat {
	private BorderStyle _border_style;
	private short _foreground_color;
	private FillPatternType _fill_partern;
	private XSSFFont _font_style;
	private HorizontalAlignment _hor_align;
	private VerticalAlignment _ver_align;
	private CellStyle _cell_style;
	private List<CellRangeAddress> _cell_range;
	private List<String> _cell_range_value;
	private Sheet _sheet;

	private class BorderFormat {
		boolean left;
		boolean right;
		boolean top;
		boolean bottom;

		/**
		 * @param left
		 * @param right
		 * @param top
		 * @param bottom
		 */
		public BorderFormat(boolean left, boolean right, boolean top, boolean bottom) {
			super();
			this.left = left;
			this.right = right;
			this.top = top;
			this.bottom = bottom;
		}

	}

	BorderFormat _boder_format;

	/**
	 * 
	 */
	public ExcelFormat() {
		super();
		this._border_style = BorderStyle.THIN;
		this._foreground_color = HSSFColor.WHITE.index;
		this._fill_partern = FillPatternType.NO_FILL;
		this._font_style = null;
		this._hor_align = HorizontalAlignment.LEFT;
		this._ver_align = VerticalAlignment.TOP;
		this._cell_style = new XSSFCellStyle(new StylesTable());
		this._cell_range = new ArrayList<CellRangeAddress>();
		this._cell_range_value = new ArrayList<String>();
		this._sheet = null;
		this._boder_format = new BorderFormat(false, false, false, false);
	}

	/**
	 * @return the _border_style
	 */
	public BorderStyle get_border_style() {
		return _border_style;
	}

	/**
	 * @param _border_style
	 *            the _border_style to set
	 */
	public void set_border_style(BorderStyle _border_style) {
		this._border_style = _border_style;
	}

	/**
	 * @return the _foreground_color
	 */
	public short get_foreground_color() {
		return _foreground_color;
	}

	/**
	 * @param _foreground_color
	 *            the _foreground_color to set
	 */
	public void set_foreground_color(short _foreground_color) {
		_cell_style.setFillForegroundColor(_foreground_color);
		this._foreground_color = _foreground_color;
	}

	/**
	 * @return the _fill_partern
	 */
	public FillPatternType get_fill_partern() {
		return _fill_partern;
	}

	/**
	 * @param _fill_partern
	 *            the _fill_partern to set
	 */
	public void set_fill_partern(FillPatternType _fill_partern) {
		_cell_style.setFillPattern(_fill_partern);
		this._fill_partern = _fill_partern;
	}

	/**
	 * @return the _font_style
	 */
	public Font get_font_style() {
		return _font_style;
	}

	/**
	 * @param _font_style
	 *            the _font_style to set
	 */
	public void set_font_style(XSSFFont _font_style) {
		_font_style.setFontName("Arial");
		_cell_style.setFont(_font_style);
		this._font_style = _font_style;
	}

	/**
	 * @return the _hor_align
	 */
	public HorizontalAlignment get_hor_align() {
		return _hor_align;
	}

	/**
	 * @param _hor_align
	 *            the _hor_align to set
	 */
	public void set_hor_align(HorizontalAlignment _hor_align) {
		_cell_style.setAlignment(_hor_align);
		this._hor_align = _hor_align;
	}

	/**
	 * @return the _ver_align
	 */
	public VerticalAlignment get_ver_align() {
		return _ver_align;
	}

	/**
	 * @param _ver_align
	 *            the _ver_align to set
	 */
	public void set_ver_align(VerticalAlignment _ver_align) {
		_cell_style.setVerticalAlignment(_ver_align);
		this._ver_align = _ver_align;
	}

	/**
	 * @return the _cell_style
	 */
	public CellStyle get_cell_style() {
		return _cell_style;
	}

	/**
	 * @param _cell_style
	 *            the _cell_style to set
	 */
	public void set_cell_style(CellStyle cell_style) {
		this._cell_style = cell_style;
	}

	

	/**
	 * @return the _sheet
	 */
	public Sheet get_sheet() {
		return _sheet;
	}

	/**
	 * @param _sheet
	 *            the _sheet to set
	 */
	public void set_sheet(Sheet _sheet) {
		this._sheet = _sheet;
	}

	private void setBorderMergeCell(BorderFormat bf)
	{
		for (CellRangeAddress cellRangeAddress : _cell_range) {
			
			if (_boder_format.left) RegionUtil.setBorderLeft(_border_style, cellRangeAddress, _sheet);
			if (_boder_format.right) RegionUtil.setBorderRight(_border_style, cellRangeAddress, _sheet);
			if (_boder_format.top) RegionUtil.setBorderTop(_border_style, cellRangeAddress, _sheet);
			if (_boder_format.bottom) RegionUtil.setBorderBottom(_border_style, cellRangeAddress, _sheet);
			
		}
	}
	private void setBorderCell(BorderFormat bf) {
		if (_boder_format.left) {
			_cell_style.setBorderLeft(_border_style);
		}
		if (_boder_format.right) {
			_cell_style.setBorderRight(_border_style);
		}
		if (_boder_format.top) {
			_cell_style.setBorderTop(_border_style);
		}
		if (_boder_format.bottom) {
			_cell_style.setBorderBottom(_border_style);
		}

	}

	public void setBorderCell(boolean left, boolean right, boolean top, boolean bottom) {
		this._boder_format = new BorderFormat(left, right, top, bottom);
		setBorderCell(_boder_format);
		setBorderMergeCell(_boder_format);
	}

	public void setBorderCell(boolean left, boolean right, boolean top, boolean bottom, BorderStyle style) {
		this._border_style = style;
		setBorderCell(left, right, top, bottom);
	}

	public CellRangeAddress mergeCell(int firstRow, int lastRow, int firstCol, int lastCol) {
		CellRangeAddress cell_range = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
		_cell_range.add(cell_range);
		return cell_range;
	}
	
	public CellRangeAddress mergeCell(int firstRow, int lastRow, int firstCol, int lastCol, String value) {
		CellRangeAddress cell_range = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
		_cell_range.add(cell_range);
		_cell_range_value.add(value);
		return cell_range;
	}
	
	private void addCellRangeToSheet()
	{
		int i=0;
		Row row = null;
		for (CellRangeAddress cellRangeAddress : _cell_range) {
			int row_num = cellRangeAddress.getFirstRow();
			if (_sheet.getRow(row_num)==null)
				row = _sheet.createRow(row_num);
			else 
				row = _sheet.getRow(row_num);
			_sheet.addMergedRegion(cellRangeAddress);
			int index = cellRangeAddress.getFirstColumn();
			Cell cell = row.createCell(index);
			cell.setCellValue(_cell_range_value.get(i++));
		}
	}
	public void refresh() {
		setBorderCell(_boder_format);
		set_fill_partern(_fill_partern);
		set_font_style(_font_style);
		set_foreground_color(_foreground_color);
		set_hor_align(_hor_align);
		set_ver_align(_ver_align);
		addCellRangeToSheet();
	}

}
