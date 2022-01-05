package puhao.simple.msword.generator.example;

import org.apache.poi.ss.formula.functions.T;
import puhao.simple.msword.generator.MSWordExporter;
import puhao.simple.msword.generator.example.cars.Car;
import puhao.simple.msword.generator.example.cars.Person;
import puhao.simple.msword.generator.example.people.TestPeople;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.net.URISyntaxException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.Arrays;

public class Example2 {
    public static void main(String args[]) throws URISyntaxException, IOException, InvocationTargetException, IllegalAccessException {

        Path templatePath = Paths.get(Example1.class.getClassLoader()
                .getResource("TheFamily.docx").toURI());

        MSWordExporter exporter = new MSWordExporter(templatePath);


        TestPeople testPeople = TestPeople.getTestFamilayExampl();


        Path outputFile = Paths.get("MyFamily.docx");

        exporter.exportTo(testPeople, outputFile);

        System.out.println("File write to " + outputFile.toAbsolutePath().toString());
    }
}
