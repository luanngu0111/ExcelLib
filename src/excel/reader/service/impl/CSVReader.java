package excel.reader.service.impl;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import excel.reader.IExcelReader;
import excel.util.ExcelUtils;


public class CSVReader  implements IExcelReader{

	private ExcelUtils _configs;
	private File _file;
	private List<String[]> _content;
	public CSVReader() {
		// TODO Auto-generated constructor stub
		this._file = null;
		this._content=new ArrayList<>();
		this._configs = ExcelUtils.getInstance();
	}
	public CSVReader(String filename, ExcelUtils utils)
	{
		this._file = new File(filename);
		this._configs = utils;
		this._content=new ArrayList<>();
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
	@Override
	public List<String[]> fileReader(String filename) {
		// TODO Auto-generated method stub
		if (this._file == null)
			_file = new File(filename);
		try {
			BufferedReader br = new BufferedReader(new FileReader(this._file));
			String line = br.readLine();
			while (line != null) {
				String[] cols = line.split(_configs.get_splitter());
				this._content.add(cols);
				line = br.readLine();
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return this._content;
	}
	
}