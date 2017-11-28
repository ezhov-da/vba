package ru.ezhov.vba;

/**
 * Преобразовываем текст для VBA
 *
 * @author ezhov_da
 */
public class VbaGeneration {

    protected IPropertiesHolder propertiesHolder = PropertiesHolder.getInstance();
    protected int countLines = 10;

    protected boolean useConnectionString;
    protected boolean addHeader;
    protected boolean addExecuteStr;
    protected String textForParse;
    protected String connectionString;
    protected String nameQuery;

    public VbaGeneration(
            boolean useConnectionString,
            boolean addHeader,
            boolean addExecuteStr,
            String textForParse,
            String connectionString,
            String nameQuery) {
        this.useConnectionString = useConnectionString;
        this.addHeader = addHeader;
        this.addExecuteStr = addExecuteStr;
        this.textForParse = textForParse;
        this.connectionString = connectionString;
        this.nameQuery = nameQuery;
    }

    public String generate() {
        String result = parse();
        String resultHeader = createHeader();
        String execute = addExecute();
        return resultHeader + result + execute;
    }

    protected String parse() {
        if (textForParse == null) {
            return "";
        }
        String[] array = textForParse.split("\n");
        String text = "";
        int length = array.length;
        for (int i = 0; i < length; i++) {
            String strFromArray = array[i];
            strFromArray = strFromArray.replaceAll("\"", "\"\"");
            if (i % countLines == 0) {
                if (i + 1 == length || (i + 1) % countLines == 0) {
                    text = text + nameQuery + " = \"" + strFromArray + "\" & vbnewline\n";
                } else {
                    text = text + nameQuery + " = " + nameQuery + " & \"" + strFromArray + "\" & vbnewline & _ \n";
                }
            } else if (i + 1 == length || (i + 1) % countLines == 0) {
                text = text + "\t\"" + strFromArray + "\"" + " & vbnewline\n";
            } else {
                text = text + "\t\"" + strFromArray + "\"" + " & vbnewline & _ \n";
            }
        }
        return text.trim();
    }

    protected String createHeader() {
        if (addHeader) {
            if (useConnectionString) {
                return String.format(propertiesHolder.getHeader(), connectionString, nameQuery, nameQuery);
            } else {
                return String.format(propertiesHolder.getHeader(), "", nameQuery, nameQuery);
            }
        }
        return "";
    }

    protected String addExecute() {
        if (addExecuteStr) {
            return String.format(propertiesHolder.getExecute(), nameQuery);
        } else {
            return "";
        }
    }
}
