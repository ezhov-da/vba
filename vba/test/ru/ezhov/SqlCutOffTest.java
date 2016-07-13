package ru.ezhov;

import org.junit.Test;

/**
 *
 * @author ezhov_da
 */
public class SqlCutOffTest {

    public SqlCutOffTest() {
    }

    @Test
    public void testGetSqlText() {
        String input = "query = query & \"\"\"1\"\" sdvdaas\" & vbnewline & _ \n"
                + "	\"2\" & vbnewline & _ \n"
                + "	\"3\" & vbnewline & _ \n"
                + "	\"4\" & vbnewline & _ \n"
                + "	\"5\" & vbnewline & _ \n"
                + "	\"6\" & vbnewline & _ \n"
                + "	\"7\" & vbnewline & _ \n"
                + "	\"8\" & vbnewline & _ \n"
                + "	\"9\" & vbnewline & _ \n"
                + "	\"10\" & vbnewline\n"
                + "query = query & \"11\" & vbnewline & _ \n"
                + "	\"12\" & vbnewline & _ \n"
                + "	\"13\" & vbnewline & _ \n"
                + "	\"14\" & vbnewline & _ \n"
                + "	\"15\" & vbnewline & _ \n"
                + "	\"16\" & vbnewline & _ \n"
                + "	\"17\" & vbnewline\n"
                + "ADO.Execute \"query\"";
        String[] sMass = input.split("\n");

        SqlCutOff sqlCutOff = new SqlCutOff();
        String string = sqlCutOff.getSqlText(input);
        System.out.println(string);

        input = "\n";
        sqlCutOff = new SqlCutOff();
        string = sqlCutOff.getSqlText(input);
        System.out.println(string);

        input = "asdads\nasdasd";
        sqlCutOff = new SqlCutOff();
        string = sqlCutOff.getSqlText(input);
        System.out.println(string);
    }

}
