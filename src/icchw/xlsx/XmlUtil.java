package icchw.xlsx;
import java.util.ArrayList;
import java.util.List;

import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 *
 */

/**
 * @author kazuhiro1-wada
 *
 */
public class XmlUtil {

	/**
	 * 当該Nodeが値が空のNodeのみから構成される場合、trueを返す。
	 * @param node
	 * @return
	 */
	public static boolean isEmptyNode(Node node) {
		NodeList nodeList = node.getChildNodes();
		for (int i = 0; i < nodeList.getLength(); i++) {
			Node childNode = nodeList.item(i);
			if (!childNode.hasChildNodes()) {
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
	 * @param node
	 * @param attributeKey
	 * @return
	 */
	public static String getAttributeValue(Node node, String attributeKey) {
		NamedNodeMap attributes = node.getAttributes();
		Node attribute = attributes.getNamedItem(attributeKey);
		if (attribute == null) {
			return null;
		}
		return attribute.getNodeValue();
	}

	/**
	 * 先頭子要素を残して、子要素を削除する
	 * @param node
	 */
	public static void removeAllChildren(Node node) {
		removeAllChilderenWithoutHeader(node, 0);
	}

	/**
	 * 先頭から指定の数の子要素を残して、子要素を削除する
	 * @param node
	 * @param headerCount 残す子要素数
	 */
	public static void removeAllChilderenWithoutHeader(Node node, int remainChildCount) {
		NodeList childNodes = node.getChildNodes();
		List<Node> removeNodeList = new ArrayList<Node>();
		for (int i = remainChildCount; i < childNodes.getLength(); i++) {
			removeNodeList.add(childNodes.item(i));
		}

		for(Node childNode : removeNodeList) {
			node.removeChild(childNode);
		}

	}
}
