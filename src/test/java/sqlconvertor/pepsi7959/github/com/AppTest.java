package sqlconvertor.pepsi7959.github.com;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.Map;

import junit.framework.Test;
import junit.framework.TestCase;
import junit.framework.TestSuite;

/**
 * Unit test for simple App.
 */
public class AppTest 
    extends TestCase
{
    /**
     * Create the test case
     *
     * @param testName name of the test case
     */
    public AppTest( String testName )
    {
        super( testName );
    }

    /**
     * @return the suite of tests being tested
     */
    public static Test suite()
    {
        return new TestSuite( AppTest.class );
    }

    /**
     * Rigourous Test :-)
     */
    public void testApp()
    {
        assertTrue( true );
    }
    
    public void testGetHeader() {
    		Map<Integer,String> header = App.getHeader("/Users/narongsak.mala/Documents/GDX/database-schema.xlsx");
		String col0 = header.get(0);
		assertTrue( col0.equals("col0"));
    }
    
    public void testreadHeader() {
		Map<Integer, LinkedList<String>> headerConf = App.readHeaderConf("/Users/narongsak.mala/Documents/GDX/Test_TemplateProcedureMapping.xlsx", 1, 1, 2, 2);
		assertEquals(headerConf.size(), 384);
		assertTrue(headerConf.get(0) != null);
		assertEquals(headerConf.get(0).get(0), "00001");
		assertEquals(headerConf.get(0).get(1), "บัตรประชาชนตนเอง");
		assertTrue(headerConf.get(383) != null);
		assertEquals(headerConf.get(383).get(0), "00384");
		assertEquals(headerConf.get(383).get(1), "หนังสือกำกับสิ่งประดิษฐ์ที่ทำจากไม้หวงห้าม");
    }
//    
//    public void testconvertRowToCell() {
//		Map<Integer, LinkedList<String>> headerConf = App.readHeaderConf("/Users/narongsak.mala/Documents/GDX/Test_TemplateProcedureMapping.xlsx", 1, 1, 2, 2);
//    		ArrayList<LinkedList<String>> listofRecord =  App.convertRowToCell(headerConf, "/Users/narongsak.mala/Documents/GDX/Test_TemplateProcedureMapping.xlsx", 2, 8, 2,-1, 0, 384);	
//    		assertTrue(listofRecord != null);
//    		assertEquals(listofRecord.get(0).get(0), "รัฐวิสาหกิจ");
//    		assertEquals(listofRecord.get(0).get(8), "00028");
//    		assertEquals(listofRecord.get(0).get(9), "หนังสือรับรองนิติบุคคล");
//    }
    
    public void testcellAddressToInt() {
    		
    		assertEquals(0, App.cellAddressToInt("A"));
    		assertEquals(25, App.cellAddressToInt("Z"));
    		assertEquals(26, App.cellAddressToInt("AA"));
    		assertEquals(51, App.cellAddressToInt("AZ"));
    		assertEquals(52, App.cellAddressToInt("bA"));
    }
}
