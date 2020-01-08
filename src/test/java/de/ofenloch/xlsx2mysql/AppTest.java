package de.ofenloch.xlsx2mysql;

import junit.framework.Test;
import junit.framework.TestCase;
import junit.framework.TestSuite;

/**
 * Unit test for simple App.
 */
public class AppTest extends TestCase {
    /**
     * Create the test case
     *
     * @param testName name of the test case
     */
    public AppTest(String testName) {
        super(testName);
    }

    /**
     * @return the suite of tests being tested
     */
    public static Test suite() {
        return new TestSuite(AppTest.class);
    }

    public void test_ColumnName2Number() {
        assertEquals(1, xlsx2mysql.ColumnName2Number("A"));
        assertEquals(4, xlsx2mysql.ColumnName2Number("D"));
        assertEquals(26, xlsx2mysql.ColumnName2Number("Z"));
        assertEquals(28, xlsx2mysql.ColumnName2Number("AB"));
        assertEquals(60, xlsx2mysql.ColumnName2Number("BH"));
        assertEquals(702, xlsx2mysql.ColumnName2Number("ZZ"));
        assertEquals(703, xlsx2mysql.ColumnName2Number("AAA"));
        assertEquals(703, xlsx2mysql.ColumnName2Number("AAA"));
        assertEquals(710, xlsx2mysql.ColumnName2Number("AAH"));
        assertEquals(2000, xlsx2mysql.ColumnName2Number("BXX"));
    }

    public void test_ColumnNumber2Name() {
        assertEquals("A", xlsx2mysql.ColumnNumberToName(1));
        assertEquals("D", xlsx2mysql.ColumnNumberToName(4));
        assertEquals("Z", xlsx2mysql.ColumnNumberToName(26));
        assertEquals("AB", xlsx2mysql.ColumnNumberToName(28));
        assertEquals("BH", xlsx2mysql.ColumnNumberToName(60));
        assertEquals("ZZ", xlsx2mysql.ColumnNumberToName(702));
        assertEquals("AAA", xlsx2mysql.ColumnNumberToName(703));
        assertEquals("AAB", xlsx2mysql.ColumnNumberToName(704));
        assertEquals("AAD", xlsx2mysql.ColumnNumberToName(706));
        assertEquals("AAH", xlsx2mysql.ColumnNumberToName(710));
        assertEquals("BXX", xlsx2mysql.ColumnNumberToName(2000));
    }
}
