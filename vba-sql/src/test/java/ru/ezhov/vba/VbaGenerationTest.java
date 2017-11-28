package ru.ezhov.vba;

import org.junit.Test;

/**
 * @author ezhov_da
 */
public class VbaGenerationTest {

    public VbaGenerationTest() {
    }

    @Test
    public void testGenerate() {
        VbaGeneration vbaGeneration =
                new VbaGeneration(
                        true,
                        true,
                        true,
                        "\"1\" sdvdaas\n"
                                + "2\n"
                                + "3\n"
                                + "4\n"
                                + "5\n"
                                + "6\n"
                                + "7\n"
                                + "8\n"
                                + "9\n"
                                + "10\n"
                                + "11\n"
                                + "12\n"
                                + "13\n"
                                + "14\n"
                                + "15\n"
                                + "16\n"
                                + "17",
                        "",
                        "query");
        String string = vbaGeneration.generate();
        System.out.println(string);
    }

}
