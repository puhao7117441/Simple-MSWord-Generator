package puhao.simple.msword.generator.example.people;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class TestPeople {
    public static final TestPeople getTestFamilayExampl(){
        TestPeople p1 = new TestPeople("孙子1","男", 23);
        TestPeople p2 = new TestPeople("孙子2","男", 18);
        TestPeople p3 = new TestPeople("孙女1","女", 16);
        TestPeople p4 = new TestPeople("孙女2","女", 30);

        TestPeople p_1 = new TestPeople("anotherSon", "Male", 50);
        TestPeople p_4 = new TestPeople("me", "Male", 48);
        TestPeople p_2 = new TestPeople("daughter1", "Female", 56);
        TestPeople p_3 = new TestPeople("daughter2", "Female", 62);
        p_1.addChild(p1);
        p_2.addChild(p2);
        p_3.addChild(p3);
        p_4.addChild(p4);


        TestPeople p__1 = new TestPeople("Dad", "Male", 87);
        TestPeople p__2 = new TestPeople("Dad's sister", "Female", 84);
        p__1.addChild(p_1);
        p__1.addChild(p_2);
        p__1.addChild(p_3);
        p__1.addChild(p_4);


        TestPeople p___1 = new TestPeople("Grandpa","Male", 109);
        p___1.addChild(p__1);
        p___1.addChild(p__2);

        return p___1;
    }


    private String name;
    private String sex;
    private int age;
    private List<TestPeople> children = new ArrayList<>();

    public TestPeople(String n, String s, int age){
        this.name = n;
        this.sex = s;
        this.age = age;
    }

    public String getHeaderName(){
        return "This is header:" + this.name;
    }
    public String getFooterUnicodeText(){
        return "这里是页脚:" + this.name;
    }

    public void addChild(TestPeople p){
        this.children.add(p);
    }
    public void addChild(String n, String s, int age){
        TestPeople c = new TestPeople(n,s,age);
        this.children.add(c);
    }

    /**
     * There do not have 'if-else' support in the template, we can use a for loop to
     * achieve the same behavior.
     * For needs like:
     * - if people no child, show 'This people no child' in the document, else show child details
     *
     * We can achieve like below:
     * --------------------------------------------
     * ${{for flag of noChildFlag}}
     *      This people no child
     * ${{end}}
     * ${{for c of children}}
     *      Name: ${c.name}
     * ${{end}}
     * --------------------------------------------
     * When there has no child, the getNoChildFlag() will return one array contain one item
     * so text 'This people no child' will repeat once in the document. And since getChildren()
     * return an empty list, the child details will repeat zero times which equals not show
     * in the document.
     *
     * When this people has child, the getNoChildFlag() will return an empty array which
     * cause text 'This people no child' not appear in the final document.
     *
     * @return
     */
    public List<String> getNoChildFlag(){
        if(this.children.size() == 0){
            return Arrays.asList("just a flag");
        }else{
            return new ArrayList<>(0);
        }
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public List<TestPeople> getChildren() {
        return children;
    }

    public int getChildAmount() {
        return children.size();
    }

    public void setChildren(List<TestPeople> children) {
        this.children = children;
    }

    @Override
    public String toString() {
        return name + ", " + sex + ", " + age + "岁, "+(children.size() == 0 ? ("没有孩子"):("有"+children.size()+"个孩子"));
    }
}
