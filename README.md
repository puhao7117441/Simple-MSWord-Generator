# Simple-MSWord-Generator
Use Apache POI to generator Microsoft Word based on given template and java POJO object
Since the Word document is very complex for me, this project only provide very simple and very limited function. 


# All below example can be found from `src/test` folder

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


  
# Complex example
In this example, we have a table which need repeat based on the data provided, also text and cell has some simple styles. Additionally the doc template contains header and footer. Like below:
![image](https://user-images.githubusercontent.com/19635360/148193169-bfddb66e-2437-4333-bb58-7e611af714e5.png)

The POJO, more details refer to `src/test/java/puhao/simple/msword/generator/example/people/TestPeople.java`
```java
public class TestPeople {
    private String name;
    private String sex;
    private int age;
    private List<TestPeople> children = new ArrayList<>();
	
	// normal getter method omitted here, except below special one
	
	// Property is not mandatory, every variable in template will 
	// map to a getter method.
	// Like in this POJO, it do not have a property with name 'headerName'
	// but we can still use ${headerName} in the template since it can 
	// find the  getHeaderName() method.
	public String getHeaderName(){return "This is header:" + this.name; }
	

    public String getFooterUnicodeText(){return "这里是页脚:" + this.name;}

	public List<String> getNoChildFlag(){
        if(this.children.size() == 0){
            return Arrays.asList("just a flag");
        }else{
            return new ArrayList<>(0);
        }
    }
	
	@Override
	public String toString() {
        return name + ", " + sex + ", " + age + "岁, "+(children.size() == 0 ? ("没有孩子"):("有"+children.size()+"个孩子"));
    }
}
```

Now, run the code
```java
TestPeople people = getPeople(); // the POJO contain the infos

MSWordExporter exporter = new MSWordExporter(templatePath);  // The template file path is the only paramter

exporter.exportTo(bob, outputFilePath);  // start generat based on the template

```


Finally, the generated doc like below:


![image](https://user-images.githubusercontent.com/19635360/148194604-8fe7f7ac-2f3e-464a-90bc-5a3f5752c7d7.png)
![image](https://user-images.githubusercontent.com/19635360/148194667-6ba4f17c-537f-48ba-9605-e1416a7474af.png)


