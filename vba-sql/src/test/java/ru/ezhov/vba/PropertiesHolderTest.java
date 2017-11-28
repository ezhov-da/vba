package ru.ezhov.vba;

import org.junit.Test;

import static org.junit.Assert.*;

public class PropertiesHolderTest {
    @Test
    public void getDefaultHeader() throws Exception {
        PropertiesHolder propertiesHolder = PropertiesHolder.getInstance();
        assertNotNull(propertiesHolder.getHeader());
    }

    @Test
    public void getDefaultExecute() throws Exception {
        PropertiesHolder propertiesHolder = PropertiesHolder.getInstance();
        assertNotNull(propertiesHolder.getExecute());
    }

}