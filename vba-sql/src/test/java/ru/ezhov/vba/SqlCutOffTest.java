package ru.ezhov.vba;

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
        
        input = test;
        sqlCutOff = new SqlCutOff();
        string = sqlCutOff.getSqlText(input);
        System.out.println(string);
    }

    private String test = "Dim ADO\n"
            + "Dim connectString\n"
            + "Set ADO = CreateObject(\"ADODB.Connection\")\n"
            + "connectString = \"\"\n"
            + "ADO.ConnectionTimeout = 0\n"
            + "ADO.CommandTimeout = 0\n"
            + "ADO.Open connectString\n"
            + "query = query & \"package ru.ezhov.vba;\" & vbnewline & _ \n"
            + "	\"\" & vbnewline & _ \n"
            + "	\"import javax.annotation.PostConstruct;\" & vbnewline & _ \n"
            + "	\"import javax.enterprise.context.RequestScoped;\" & vbnewline & _ \n"
            + "	\"import javax.inject.Named;\" & vbnewline & _ \n"
            + "	\"import ru.ezhov.SqlCutOff;\" & vbnewline & _ \n"
            + "	\"import ru.ezhov.VbaGeneration;\" & vbnewline & _ \n"
            + "	\"\" & vbnewline & _ \n"
            + "	\"/**\" & vbnewline & _ \n"
            + "	\" *\" & vbnewline\n"
            + "query = query & \" * @author ezhov_da\" & vbnewline & _ \n"
            + "	\" */\" & vbnewline & _ \n"
            + "	\"@RequestScoped\" & vbnewline & _ \n"
            + "	\"@Named(\"\"generator\"\")\" & vbnewline & _ \n"
            + "	\"public class GeneratorSqlToExcel {\" & vbnewline & _ \n"
            + "	\"\" & vbnewline & _ \n"
            + "	\"    private String strConnection;\" & vbnewline & _ \n"
            + "	\"    private String tempQuery;\" & vbnewline & _ \n"
            + "	\"    private String code;\" & vbnewline & _ \n"
            + "	\"    private String finalCode;\" & vbnewline\n"
            + "query = query & \"    private boolean useConnectionString;\" & vbnewline & _ \n"
            + "	\"    private boolean addHeaderConnection;\" & vbnewline & _ \n"
            + "	\"    private boolean addConnectionString;\" & vbnewline & _ \n"
            + "	\"    private String codeVbaToSql;\" & vbnewline & _ \n"
            + "	\"    private String finalCodeToSql;\" & vbnewline & _ \n"
            + "	\"\" & vbnewline & _ \n"
            + "	\"    @PostConstruct\" & vbnewline & _ \n"
            + "	\"    public void init() {\" & vbnewline & _ \n"
            + "	\"        tempQuery = \"\"query\"\";\" & vbnewline & _ \n"
            + "	\"    }\" & vbnewline\n"
            + "query = query & \"\" & vbnewline & _ \n"
            + "	\"    public String getStrConnection() {\" & vbnewline & _ \n"
            + "	\"        return strConnection;\" & vbnewline & _ \n"
            + "	\"    }\" & vbnewline & _ \n"
            + "	\"\" & vbnewline & _ \n"
            + "	\"    public void setStrConnection(String strConnection) {\" & vbnewline & _ \n"
            + "	\"        this.strConnection = strConnection;\" & vbnewline & _ \n"
            + "	\"    }\" & vbnewline & _ \n"
            + "	\"\" & vbnewline & _ \n"
            + "	\"    public String getTempQuery() {\" & vbnewline\n"
            + "query = query & \"        return tempQuery;\" & vbnewline & _ \n"
            + "	\"    }\" & vbnewline & _ \n"
            + "	\"\" & vbnewline & _ \n"
            + "	\"    public void setTempQuery(String tempQuery) {\" & vbnewline & _ \n"
            + "	\"        this.tempQuery = tempQuery;\" & vbnewline & _ \n"
            + "	\"    }\" & vbnewline & _ \n"
            + "	\"\" & vbnewline & _ \n"
            + "	\"    public String getCode() {\" & vbnewline & _ \n"
            + "	\"        return code;\" & vbnewline & _ \n"
            + "	\"    }\" & vbnewline\n"
            + "query = query & \"\" & vbnewline & _ \n"
            + "	\"    public void setCode(String code) {\" & vbnewline & _ \n"
            + "	\"        this.code = code;\" & vbnewline & _ \n"
            + "	\"    }\" & vbnewline & _ \n"
            + "	\"\" & vbnewline & _ \n"
            + "	\"    public String getFinalCode() {\" & vbnewline & _ \n"
            + "	\"        VbaGeneration generation = new VbaGeneration(useConnectionString, addHeaderConnection, addConnectionString, code, strConnection, tempQuery);\" & vbnewline & _ \n"
            + "	\"        String result = generation.generate();\" & vbnewline & _ \n"
            + "	\"        return result;\" & vbnewline & _ \n"
            + "	\"    }\" & vbnewline\n"
            + "query = query & \"\" & vbnewline & _ \n"
            + "	\"    public void setFinalCode(String finalCode) {\" & vbnewline & _ \n"
            + "	\"        this.finalCode = finalCode;\" & vbnewline & _ \n"
            + "	\"    }\" & vbnewline & _ \n"
            + "	\"\" & vbnewline & _ \n"
            + "	\"    public boolean isUseConnectionString() {\" & vbnewline & _ \n"
            + "	\"        return useConnectionString;\" & vbnewline & _ \n"
            + "	\"    }\" & vbnewline & _ \n"
            + "	\"\" & vbnewline & _ \n"
            + "	\"    public void setUseConnectionString(boolean useConnectionString) {\" & vbnewline\n"
            + "query = query & \"        this.useConnectionString = useConnectionString;\" & vbnewline & _ \n"
            + "	\"    }\" & vbnewline & _ \n"
            + "	\"\" & vbnewline & _ \n"
            + "	\"    public boolean isAddHeaderConnection() {\" & vbnewline & _ \n"
            + "	\"        return addHeaderConnection;\" & vbnewline & _ \n"
            + "	\"    }\" & vbnewline & _ \n"
            + "	\"\" & vbnewline & _ \n"
            + "	\"    public void setAddHeaderConnection(boolean addHeaderConnection) {\" & vbnewline & _ \n"
            + "	\"        this.addHeaderConnection = addHeaderConnection;\" & vbnewline & _ \n"
            + "	\"    }\" & vbnewline\n"
            + "query = query & \"\" & vbnewline & _ \n"
            + "	\"    public boolean isAddConnectionString() {\" & vbnewline & _ \n"
            + "	\"        return addConnectionString;\" & vbnewline & _ \n"
            + "	\"    }\" & vbnewline & _ \n"
            + "	\"\" & vbnewline & _ \n"
            + "	\"    public void setAddConnectionString(boolean addConnectionStrong) {\" & vbnewline & _ \n"
            + "	\"        this.addConnectionString = addConnectionStrong;\" & vbnewline & _ \n"
            + "	\"    }\" & vbnewline & _ \n"
            + "	\"\" & vbnewline & _ \n"
            + "	\"    public String getFinalCodeToSql() {\" & vbnewline\n"
            + "query = query & \"        SqlCutOff sqlCutOff = new SqlCutOff();\" & vbnewline & _ \n"
            + "	\"        return \"\"<pre>\"\" + sqlCutOff.getSqlText(codeVbaToSql) + \"\"</pre>\"\";\" & vbnewline & _ \n"
            + "	\"    }\" & vbnewline & _ \n"
            + "	\"\" & vbnewline & _ \n"
            + "	\"    public String getCodeVbaToSql() {\" & vbnewline & _ \n"
            + "	\"        return codeVbaToSql;\" & vbnewline & _ \n"
            + "	\"    }\" & vbnewline & _ \n"
            + "	\"\" & vbnewline & _ \n"
            + "	\"    public void setCodeVbaToSql(String codeVbaToSql) {\" & vbnewline & _ \n"
            + "	\"        this.codeVbaToSql = codeVbaToSql;\" & vbnewline\n"
            + "query = query & \"    }\" & vbnewline & _ \n"
            + "	\"\" & vbnewline & _ \n"
            + "	\"}\" & vbnewline\n"
            + "ADO.Execute \"query\"";
}
