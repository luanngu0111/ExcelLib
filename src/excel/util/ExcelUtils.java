package excel.util;

import java.util.ArrayList;
import java.util.List;

public class ExcelUtils {

	private int _sheet_id;
	private String _sheet_name;
	private String _splitter = ";";
	private List<String> _header;
	private ExcelFormat _format;
	private static ExcelUtils _instance;
	private boolean _overwrite = true;

	private ExcelUtils() {
		// TODO Auto-generated constructor stub
		_sheet_id = -1;
		_sheet_name = "Sheet 1";
		_header = new ArrayList<>();
		_format = new ExcelFormat();
	}

	public static ExcelUtils getNewInstance() {
		return new ExcelUtils();
	}

	public static ExcelUtils getInstance() {
		if (_instance == null) {
			_instance = new ExcelUtils();
		}
		return _instance;
	}

	/**
	 * @return the _sheet_id
	 */
	public int get_sheet_id() {
		return _sheet_id;
	}

	/**
	 * @param _sheet_id
	 *            the _sheet_id to set
	 */
	public void set_sheet_id(int _sheet_id) {
		this._sheet_id = _sheet_id;
	}

	/**
	 * @return the _sheet_name
	 */
	public String get_sheet_name() {
		return _sheet_name;
	}

	/**
	 * @param _sheet_name
	 *            the _sheet_name to set
	 */
	public void set_sheet_name(String _sheet_name) {
		this._sheet_name = _sheet_name;
	}

	/**
	 * @return the _splitter
	 */
	public String get_splitter() {
		return _splitter;
	}

	/**
	 * @param _splitter
	 *            the _splitter to set
	 */
	public void set_splitter(String _splitter) {
		this._splitter = _splitter;
	}

	/**
	 * @return the _header
	 */
	public List<String> get_header() {
		return _header;
	}

	/**
	 * @param _header
	 *            the _header to set
	 */
	public void set_header(List<String> _header) {
		this._header = _header;
	}

	/**
	 * @return the _format
	 */
	public ExcelFormat get_format() {
		return _format;
	}

	/**
	 * @param _format
	 *            the _format to set
	 */
	public void set_format(ExcelFormat _format) {
		this._format = _format;
	}

	public boolean isOverwrite() {
		return _overwrite;
	}

	public void setOverwrite(boolean overwrite) {
		this._overwrite = overwrite;
	}

}
