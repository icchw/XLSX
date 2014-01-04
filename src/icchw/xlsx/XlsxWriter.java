package icchw.xlsx;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellReference;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;


public class XlsxWriter {


	/**
	 * sheetのxmlにデータを追加する
	 * header行があれば残して、後続に追加する
	 * header行の次の行の書式をコピーして使用する
	 * header行の次の行に式が入力されている場合、各行にコピーする
	 */
	@SuppressWarnings("unchecked")
	public static void addDataToSheet(Document sheetXml, List<? extends XlsxWritable> datas) {
		Node sheetDataNode = sheetXml.getDocumentElement().getElementsByTagName("sheetData").item(0);
		int startRowNumber = getStartRowNumber(sheetDataNode);

		Node startRowNode = getRowNode(sheetXml, sheetDataNode, startRowNumber);

		// 先頭（header除く）行に対する書式
		Map<Integer, Integer> styleMap = getStyleMap(startRowNode);

		// 列全体に対する書式（範囲設定）
		Map<Integer, Integer> colStyleMap = getColStyleMap(sheetXml.getDocumentElement().getElementsByTagName("cols").item(0));

		// 先頭（header除く）行に対する数式
		Map<Integer, String> functionMap = getFunctionMap(startRowNode);

		// header行を残して全行を一度削除
		XmlUtil.removeAllChilderenWithoutHeader(sheetDataNode, startRowNumber);

		for (int row = 0; row < datas.size(); row++) {
			// 行の作成
			Node rowNode  = createRowNode(sheetXml, sheetDataNode, startRowNumber + row);

			// 各列の書込み
			Map<Integer, Object> map = datas.get(row).getMap();
			for (Integer col : getKeyList(map.keySet(), functionMap.keySet())) {
				Node cellNode = createCellNode(sheetXml, startRowNumber + row, col, map.get(col), styleMap.get(col), colStyleMap.get(col));
				// 式のコピー
				cellNode = addFuntion(sheetXml, cellNode, FunctionUtil.convertCellReferencesRow(functionMap.get(col), startRowNumber, startRowNumber + row), startRowNumber +row, col);
				if (cellNode != null) {
					rowNode.appendChild(cellNode);
				}
			}
			sheetDataNode.appendChild(rowNode);
		}
	}

	/**
	 * Cellにfunctionを追加する
	 * @param sheetXml
	 * @param cellNode nullの場合、Cell Nodeをcreateして式を追加する
	 * @param functionStr nullまたは空の場合、引数をそのまま返す
	 * @param row
	 * @param col
	 * @return
	 */
	private static Node addFuntion(Document sheetXml, Node cellNode, String functionStr, int row, int col) {
		if (functionStr == null || functionStr.isEmpty()) {
			return cellNode;
		}
		if (cellNode != null) {
			return FunctionUtil.addFunctionStr(cellNode, functionStr);
		}
		Node newCellNode = createCell(sheetXml, row, col, null, null, null);
		return FunctionUtil.addFunctionStr(newCellNode, functionStr);
	}

	/**
	 * keySetsをsort済のArrayListに変換する（XML書き込み時に左列から順に処理する必要があるため）
	 * @param keySets
	 * @return
	 */
	@SafeVarargs
	private static List<Integer> getKeyList(Set<Integer> ... keySets) {
		Set<Integer> key = new TreeSet<Integer>();
		for (Set<Integer> set : keySets) {
			for (Integer integer : set) {
				key.add(integer);
			}
		}
		List<Integer> list = new ArrayList<Integer>();
		list.addAll(key);
		return list;
	}

	/**
	 * colsに定義されているstyleを、Map<列番号, style index>に変換する
	 * @param cols
	 * @return
	 */
	private static Map<Integer, Integer> getColStyleMap(Node cols) {
		Map<Integer, Integer> colStyleMap = new HashMap<Integer, Integer>();
		if (cols == null) {
			return colStyleMap;
		}

		NodeList childNodes = cols.getChildNodes();
		for (int i = 0; i < childNodes.getLength(); i++) {
			Node cellNode = childNodes.item(i);
			String minStr = XmlUtil.getAttributeValue(cellNode, "min");
			String maxStr = XmlUtil.getAttributeValue(cellNode, "max");
			String style = XmlUtil.getAttributeValue(cellNode, "style");
			if (style == null || style.isEmpty()) {
				continue;
			}
			int[] colNoArray = getColNoArray(minStr, maxStr);
			for (int colNo : colNoArray) {
				colStyleMap.put(colNo, Integer.parseInt(style));
			}
		}
		return colStyleMap;
	}

	private static int[] getColNoArray(String minStr, String maxStr) {
		int min = Integer.parseInt(minStr);
		int max = Integer.parseInt(maxStr);
		int size = max - min + 1;
		int[] result = new int[size];
		for (int i = 0; i < size; i++) {
			result[i] = min + i;
		}
		return result;
	}

	/**
	 * 指定行のStyleを、Map<列番号, style index>に変換する
	 * @param rowNode nullの場合、空のMapを返す
	 * @return
	 */
	public static Map<Integer, Integer> getStyleMap(Node rowNode) {
		Map<Integer, Integer> styleMap = new HashMap<Integer, Integer>();
		if (rowNode == null) {
			return styleMap;
		}

		NodeList childNodes = rowNode.getChildNodes();
		for (int i = 0; i < childNodes.getLength(); i++) {
			Node cellNode = childNodes.item(i);
			String r = XmlUtil.getAttributeValue(cellNode, "r");
			if (r == null || r.isEmpty()) {
				continue;
			}
			String s = XmlUtil.getAttributeValue(cellNode, "s");
			if (s == null || s.isEmpty()) {
				continue;
			}
			styleMap.put((int) getColumnIndex(r), Integer.parseInt(s));
		}
		return styleMap;
	}

	/**
	 * 指定行の式を、Map<列番号, 式の文字列>に変換する
	 * @param rowNode nullの場合、空のMapを返す
	 * @return
	 */
	public static Map<Integer, String> getFunctionMap(Node rowNode) {
		Map<Integer, String> functionMap = new HashMap<Integer, String>();
		if (rowNode == null) {
			return functionMap;
		}

		NodeList childNodes = rowNode.getChildNodes();
		for (int i = 0; i < childNodes.getLength(); i++) {
			Node cellNode = childNodes.item(i);
			String r = XmlUtil.getAttributeValue(cellNode, "r");
			String f = FunctionUtil.getFunctionStr(cellNode);
			if (f != null) {
				functionMap.put((int) getColumnIndex(r), f);
			}
		}
		return functionMap;
	}

	/**
	 * rowを取得する、存在しない場合はnullを返す
	 * @param sheetXml
	 * @param sheetDataNode
	 * @param rowNumber
	 * @return
	 */
	public static Node getRowNode(Document sheetXml, Node sheetDataNode, int rowNumber) {
		NodeList rows = sheetDataNode.getChildNodes();
		for (int i = 0; i < rows.getLength(); i++) {
			Node node = rows.item(i);
			String rValue = XmlUtil.getAttributeValue(node, "r");
			if (rValue == null) {
				continue;
			}
			if (rValue.equals(Integer.toString(rowNumber + 1))) {
				return node;
			}
		}
		return null;
	}

	/**
	 * rowを作成する
	 * @param sheetXml
	 * @param sheetDataNode
	 * @param rowNumber
	 * @return
	 */
	public static Node createRowNode(Document sheetXml, Node sheetDataNode, int rowNumber) {
		Element newNode = sheetXml.createElement("row");
		newNode.setAttribute("r", Integer.toString(rowNumber + 1));
		return newNode;
	}

	/**
	 * 全列が空または式の行を最初の行と判断する
	 * 空のシートの場合は、0を返す
	 * @param sheetDataNode
	 * @return
	 */
	public static int getStartRowNumber(Node sheetDataNode) {
		NodeList rows = sheetDataNode.getChildNodes();
		for(int i = 0; i < rows.getLength(); i++) {
			if (FunctionUtil.isEmptyNode(rows.item(i))) {
				return i;
			}
		}
		return rows.getLength();
	}


	/**
	 * (0, 0)⇒(A1)
	 * @param rowIndex
	 * @param columnIndex
	 * @return
	 */
	private static String getCellName(int rowIndex, int columnIndex) {
		return new CellReference(rowIndex, columnIndex).formatAsString();
	}

	/**
	 * (A2)⇒(1)
	 * @param cellReference
	 * @return
	 */
	private static short getColumnIndex(String cellReference) {
		return new CellReference(cellReference).getCol();
	}


	/**
	 * @param row
	 * @param col
	 * @param value
	 * @param styleIndex
	 * @param styleIndexCol
	 * @return
	 */
	public static Node createCellNode(Document sheetXml, int row, int col, Object value, Integer styleIndex, Integer styleIndexCol) {
		if (value == null) {
			if (styleIndex == null) {
				return null;
			} else {
				return createCell(sheetXml, row, col, null, styleIndex, styleIndexCol);
			}
		}
		if (value instanceof String) {
			return createStringCellNode(sheetXml, row, col, (String) value, styleIndex, styleIndexCol);
		} else if (value instanceof Number) {
			return createNumberCellNode(sheetXml, row, col, (Number) value, styleIndex, styleIndexCol);
		} else if (value instanceof BigDecimal) {
			return createNumberCellNode(sheetXml, row, col, (BigDecimal) value, styleIndex, styleIndexCol);
		} else if (value instanceof Date) {
			return createDateCellNode(sheetXml, row, col, (Date) value, styleIndex, styleIndexCol);
		} else if (value instanceof Calendar) {
			return createDateCellNode(sheetXml, row, col, ((Calendar) value).getTime(), styleIndex, styleIndexCol);
		} else {
			return createStringCellNode(sheetXml, row, col, value.toString(), styleIndex, styleIndexCol);
		}
	}

	private static Element createCell(Document sheetXml, int row, int col, String attributeT, Integer styleIndex, Integer styleIndexCol) {
		Element colNode = sheetXml.createElement("c");
		colNode.setAttribute("r", getCellName(row, col));
		if (attributeT != null && !attributeT.isEmpty()) {
			colNode.setAttribute("t", attributeT);
		}
		if (styleIndex != null) {
			colNode.setAttribute("s", styleIndex.toString());
		} else if (styleIndexCol != null) {
			colNode.setAttribute("s", styleIndexCol.toString());
		}
		return colNode;
	}

	private static Node createStringCellNode(Document sheetXml, int row, int col, String value, Integer styleIndex, Integer styleIndexCol) {
		Element colNode = createCell(sheetXml, row, col, "inlineStr", styleIndex, styleIndexCol);

		Element isNode = sheetXml.createElement("is");
		colNode.appendChild(isNode);

		Element tNode = sheetXml.createElement("t");
		isNode.appendChild(tNode);
		tNode.appendChild(sheetXml.createTextNode(value));

		return colNode;
	}

	private static Node createNumberCellNode(Document sheetXml, int row, int col, Number value, Integer styleIndex, Integer styleIndexCol) {
		Element colNode = createCell(sheetXml, row, col, "n", styleIndex, styleIndexCol);

		Element vNode = sheetXml.createElement("v");
		colNode.appendChild(vNode);

		vNode.appendChild(sheetXml.createTextNode(Double.toString(value.doubleValue())));

		return colNode;
	}

	private static Node createNumberCellNode(Document sheetXml, int row, int col, BigDecimal value, Integer styleIndex, Integer styleIndexCol) {
		return createNumberCellNode(sheetXml, row, col, value.doubleValue(), styleIndex, styleIndexCol);
	}

	private static Node createDateCellNode(Document sheetXml, int row, int col, Date value, Integer styleIndex, Integer styleIndexCol) {
		Element colNode = createCell(sheetXml, row, col, "n", styleIndex, styleIndexCol);

		Element vNode = sheetXml.createElement("v");
		colNode.appendChild(vNode);

		vNode.appendChild(sheetXml.createTextNode(Double.toString(DateUtil.getExcelDate(value))));

		return colNode;
	}
}
