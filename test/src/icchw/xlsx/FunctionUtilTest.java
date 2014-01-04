package icchw.xlsx;
		import static org.junit.Assert.*;
import icchw.xlsx.FunctionUtil;

import org.junit.Test;


public class FunctionUtilTest extends FunctionUtil {

	@Test
	public void testConvertCellReferences() {
		assertEquals("VLOOKUP(Sheet1!D2,メールアドレス!$C$2:$D$41,2,FALSE)", convertCellReferencesRow("VLOOKUP(Sheet1!D2,メールアドレス!$C$2:$D$41,2,FALSE)", 1, 1));
		assertEquals("VLOOKUP(Sheet1!D3,メールアドレス!$C$2:$D42,2,FALSE)", convertCellReferencesRow("VLOOKUP(Sheet1!D2,メールアドレス!$C$2:$D41,2,FALSE)", 1, 2));
		assertEquals("=A2*2", convertCellReferencesRow("=A1*2", 1, 2));
		assertEquals("=A2+B2", convertCellReferencesRow("=A1+B1", 1, 2));
		assertEquals("=(A2)*2", convertCellReferencesRow("=(A1)*2", 1, 2));
	}

	@Test
	public void testConvertCellReferencesWithStr() {
		// ="A1:"&A1
		assertEquals("=\"A1:\"\"\" & A2", convertCellReferencesRow("=\"A1:\"\"\" & A1", 1, 2));
		// =A1&", A1:"&A1
		assertEquals("=A2&\", A1:\"\"\" &A2", convertCellReferencesRow("=A1&\", A1:\"\"\" &A1", 1, 2));
	}

	@Test
	public void testConverCellReferencesRange() {
		assertEquals("=AA:AA", convertCellReferences("=Z:Z", 1, 1, 2, 2));
	}

	@Test
	public void testConverCellReferencesRange2() {
		assertEquals("=12:12", convertCellReferencesRow("=11:11", 1, 2));
	}

	@Test
	public void testConverCellReferencesRange3() {
		assertEquals("=SUM(E14:F17)", convertCellReferences("=SUM(D13:E16)", 10, 10, 11, 11));
	}


	@Test
	public void TestConvertCellReference1() {
		assertEquals("AA100", convertCellReference("Z99", 1, 1, 2, 2));
	}

	@Test
	public void TestConvertCellReference2() {
		assertEquals("$AA99", convertCellReference("$AA100", 2, 2, 1, 1));
	}

	@Test
	public void TestConvertCellReference3() {
		assertEquals("Z$100", convertCellReference("AA$100", 2, 2, 1, 1));
	}

	@Test
	public void TestConvertCellReferenceRange() {
		assertEquals("AA:AAA", convertCellReference("Z:ZZ", 1, 1, 2, 1));
	}

	@Test
	public void TestConvertCellReferenceRange2() {
		assertEquals("$Z:$ZZ", convertCellReference("$Z:$ZZ", 1, 1, 2, 1));
	}

}
