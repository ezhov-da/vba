package ru.ezhov.vba;

/**
 * Класс, который очищает запрос SQL от кода vba
 *
 * @author ezhov_da
 */
public class SqlCutOff {
    public String getSqlText(String text) {
        if (text == null || "".equals(text)) {
            return "";
        }
        String[] sMass = text.split("\n");
        StringBuilder builder = new StringBuilder();
        for (String st : sMass) {
            int firstIndex = st.indexOf("\"");

            if (firstIndex == -1) {
                return text;
            }
            String first = st.substring(firstIndex + 1, st.length());
            int lastIndex = first.lastIndexOf("\"");
            if (lastIndex == -1) {
                return text;
            }
            String last = first.substring(0, lastIndex);
            String finalRow = last.replaceAll("\"\"", "\"");
            builder.append(finalRow);
            builder.append("\n");
        }
        return builder.toString();
    }
}
