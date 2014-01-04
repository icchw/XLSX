package icchw.xlsx;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.util.CellReference;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 * Excel式のコピー用のUtilクラス
 */
public class FunctionUtil {

	private static final String REG_CELL_REFERENCE = "\\$?[A-Z]+\\$?[0-9]+|\\$?[A-Z]+:\\$?[A-Z]+|\\$?[0-9]+:\\$?[0-9]+";
	private static final Pattern PATTERN_CELL_REFERENCE = Pattern.compile(REG_CELL_REFERENCE);
	private static final String REG_STR = "\".*\"";
	private static final Pattern PATTERN_STR = Pattern.compile(REG_STR);


	/**
	 * 式を含むセルの場合、式を返す
	 * それ以外の場合は、nullを返す
	 * @param cellNode
	 * @return
	 */
	public static String getFunctionStr(Node cellNode) {
		NodeList nodeList = cellNode.getChildNodes();
		for (int i = 0; i < nodeList.getLength(); i++) {
			Node n = nodeList.item(i);
			if (n.getNodeName().equals("f")) {
				return n.getChildNodes().item(0).getTextContent();
			}
		}
		return null;
	}

	/**
	 * セルに式を追加する
	 * @param cellNode
	 * @param functionStr
	 */
	public static Node addFunctionStr(Node cellNode, String functionStr) {
		Document document = cellNode.getOwnerDocument();
		Node fNode = document.createElement("f");
		Node textNode = document.createTextNode(functionStr);
		fNode.appendChild(textNode);
		cellNode.insertBefore(fNode, cellNode.getFirstChild());
		return cellNode;
	}


	/**
	 * 当該Nodeが値が空または式のみを含むNodeのみから構成される場合、trueを返す。
	 * @param node
	 * @return
	 */
	public static boolean isEmptyNode(Node node) {
		NodeList nodeList = node.getChildNodes();
		for (int i = 0; i < nodeList.getLength(); i++) {
			Node childNode = nodeList.item(i);
			if (getFunctionStr(childNode) != null) {
				continue;
			} else if (!childNode.hasChildNodes()) {
				String value = childNode.getNodeValue();
				if (value != null && !value.isEmpty()) {
					return false;
				}
			} else {
				if (!isEmptyNode(childNode)) {
					return false;
				}
			}
		}
		return true;
	}

	/**
	 * 式のセル参照を変換する（行の移動）<br/>
	 * {@link #convertCellReferences(String, int, int, int, int)}を呼ぶ
	 * @param originalStr nullの場合、nullを返す
	 * @param srcRow セル参照取得セルの列番号
	 * @param destRow 設定先セルの列番号
	 * @return
	 */
	public static String convertCellReferencesRow(String originalStr, int srcRow, int destRow) {
		return convertCellReferences(originalStr, 0, srcRow, 0, destRow);
	}

	/**
	 * 式のセル参照を変換する
	 * @param originalStr nullの場合、nullを返す
	 * @param srcCol
	 * @param srcRow
	 * @param destCol
	 * @param destRow
	 * @return
	 */
	public static String convertCellReferences(String originalStr, int srcCol, int srcRow, int destCol, int destRow) {
		if (originalStr == null || originalStr.isEmpty()) {
			return originalStr;
		}
		String resultStr = "";
		// 文字列を前方から順に処理
		while(!originalStr.isEmpty()) {
			Matcher strMatcher = PATTERN_STR.matcher(originalStr);
			Matcher referenceMatcher = PATTERN_CELL_REFERENCE.matcher(originalStr);
			if (!referenceMatcher.find()) {
				// セル参照を含まない場合、そのまま返す
				resultStr += originalStr;
				return resultStr;
			} else if (!strMatcher.find() || (referenceMatcher.start() < strMatcher.start())) {
				// セル参照を含むが文字列を含まない場合、もしくは、セル参照が文字列より前にある場合、最初のセル参照を変換して処理を続ける
				int referenceStart = referenceMatcher.start();
				resultStr += originalStr.substring(0, referenceStart);
				originalStr = originalStr.substring(referenceStart);
				resultStr += convertCellReference(referenceMatcher.group(), srcCol, srcRow, destCol, destRow);
				originalStr = originalStr.replaceFirst(REG_CELL_REFERENCE, "");
			} else {
				// 文字列がセル参照より前にある場合、文字列を移して処理を続ける
				resultStr += originalStr.substring(0, strMatcher.end(0));
				originalStr = originalStr.substring(strMatcher.end(0));
			}
		}
		return resultStr;

	}

	/**
	 * コピー元セルからコピー先セルにセル参照をコピーした場合の、コピー後の参照を返す
	 * ex. ("A1", 1, 1, 2, 2) => "B2"
	 * @param cellReference
	 * @param srcCol
	 * @param srcRow
	 * @param destCol
	 * @param destRow
	 * @return
	 */
	public static String convertCellReference(String cellReference, int srcCol, int srcRow, int destCol, int destRow) {
		if (cellReference.contains(":")) {
			// 範囲の場合、再帰的に処理
			String[] range = cellReference.split(":");
			return convertCellReference(range[0], srcCol, srcRow, destCol, destRow) + ":" + convertCellReference(range[1], srcCol, srcRow, destCol, destRow);
		}
		CellReference cr = new CellReference(cellReference);
		String col;
		if (cr.getCol() == -1 ) {
			col = "";
		} else if (cr.isColAbsolute()) {
			col = "$" + CellReference.convertNumToColString(cr.getCol());
		} else {
			col = CellReference.convertNumToColString(cr.getCol() - srcCol + destCol);
		}

		String row;
		if (cr.getRow() == -1) {
			row = "";
		} else if (cr.isRowAbsolute()) {
			row = "$" + (cr.getRow() + 1);
		} else {
			row = Integer.toString(cr.getRow() - srcRow + destRow + 1);
		}
		return col + row;
	}
}
