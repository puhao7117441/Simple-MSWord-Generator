package puhao.simple.msword.generator.example.cars;


import puhao.simple.msword.generator.MSWordExporter;
import puhao.simple.msword.generator.example.cars.pojo.Car;
import puhao.simple.msword.generator.example.cars.pojo.Person;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.net.URISyntaxException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Arrays;

public class Example1 {
    public static void main(String args[]) throws URISyntaxException, IOException, InvocationTargetException, IllegalAccessException {

        Path templatePath = Paths.get(Example1.class.getClassLoader()
                .getResource("TheCars.docx").toURI());

        MSWordExporter exporter = new MSWordExporter(templatePath);


        Car car1 = new Car("Model 3", "$23,000", LocalDate.of(2018,3,2));
        Car car2 = new Car("Stone Z/2011", "$132,210", LocalDate.of(2011,1,21));
        Car car3 = new Car("Unicode唐", "￥32,210", LocalDate.of(2020,5,11));

        Person bob = new Person();
        bob.setName("Bob");
        bob.setCars(Arrays.asList(car1, car2, car3));

        Path outputFile = Paths.get("Bob_sCars.docx");

        exporter.exportTo(bob, outputFile);

        System.out.println("File write to " + outputFile.toAbsolutePath().toString());
    }
}
