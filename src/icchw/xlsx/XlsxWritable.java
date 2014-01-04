package icchw.xlsx;
import java.util.Map;


public interface XlsxWritable {

	/**
	 * 書き込む列のindex（0始まり）に対する値をMapとして返す<br/>
	 * String=>String, Number,BigDecimal,Date,Calender=>Number, others=>toStringしてセルに書き込む
	 * @return
	 */
	Map<Integer, Object> getMap();
}
