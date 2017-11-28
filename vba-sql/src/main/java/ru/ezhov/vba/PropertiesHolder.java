package ru.ezhov.vba;

import java.io.IOException;
import java.util.Properties;

public class PropertiesHolder implements IPropertiesHolder {

    private static PropertiesHolder propertiesHolder;
    private static Properties properties;

    private PropertiesHolder() {
    }

    public static PropertiesHolder getInstance() {
        if (propertiesHolder == null) {
            propertiesHolder = new PropertiesHolder();

            Properties properties = new Properties();
            try {
                properties.load(PropertiesHolder.class.getResourceAsStream("/vba.properties"));
            } catch (IOException e) {
                e.printStackTrace();
            }

            propertiesHolder.properties = properties;
        }

        return propertiesHolder;
    }

    public String getHeader() {
        return properties.getProperty("default.header");
    }

    public String getExecute() {
        return properties.getProperty("default.execute");
    }

}
