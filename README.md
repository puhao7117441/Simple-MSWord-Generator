# Simple-MSWord-Generator
Use Apache POI to generator Microsoft Word based on given template and java POJO object
Since the Word document is very complex for me, this project only provide very simple and very limited function. 


# Simple Example
First create your Microsoft Word template, must be a docx file, like below.
![image](https://user-images.githubusercontent.com/19635360/148183208-e8d727cf-b716-4633-8e5c-6e43f147ed79.png)

Then creat the java POJO object
```java
  public class Car {
    private String model;
    private String price;
    private LocalDate boughtAt;

    // Mandatory Getter method omitted here
  }
```
```java
public class Person {
    private String name;
    private List<Car> cars;
  
   // Mandatory Getter method omitted here except below one
   public Integer getTotalCount() { return cars == null ? 0 : cars.size();  }
}
```

Now, run the code
```java
Person bob = getDetailsForBob(); // the POJO contain the infos

MSWordExporter exporter = new MSWordExporter(templatePath);  // The template file path is the only paramter

exporter.exportTo(bob, outputFilePath);  // start generat based on the template

```

Finally, the generated doc like below:
![image](https://user-images.githubusercontent.com/19635360/148184788-0cc0c432-c67a-43b6-9d4a-5f154336ba4c.png)


  
