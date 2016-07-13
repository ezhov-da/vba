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
        String input = "query = \"select\" & vbnewline & _ \n"
                + "	\"    t0.id,\" & vbnewline & _ \n"
                + "	\"    t0.name\" & vbnewline & _ \n"
                + "	\"from Otz.dbo.T_E_testProstoTak t0\" & vbnewline & _ \n"
                + "	\"inner join OTZ.dbo.T_E__agaga t1 on\" & vbnewline & _ \n"
                + "	\"    t0.id = t1.id\" & vbnewline & _ \n"
                + "	\"left join OTZ.dbo.T_E__agaga t1 on\" & vbnewline & _ \n"
                + "	\"    t0.id = t1.id    \" & vbnewline";
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
