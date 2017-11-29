/**
 * 
 */
package excel.reader;

import java.util.List;

/**
 * @author LuanNgu
 *
 */
public interface IExcelReader {
	public List<String[]> fileReader(String filename);
}
