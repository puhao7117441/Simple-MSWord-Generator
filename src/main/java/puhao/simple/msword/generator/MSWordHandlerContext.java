package puhao.simple.msword.generator;


import org.apache.commons.lang3.StringUtils;
import puhao.simple.msword.generator.parse.DocElementUnit;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;

public class MSWordHandlerContext {
    public static final String ROOT_OBJECT_REFER_NAME = "_";
    private Object rootObject = null;
    private ArrayList<ContextReferObject> referObjectStack = new ArrayList<>();
    private Map<Class, Map<String, Method>> getterMethodMap = new HashMap<>();

    public MSWordHandlerContext(Object rootObject){
        this.rootObject = rootObject;
    }

    public boolean hasReferObject(String referObjName) {
        if(referObjName == null){
            throw new NullPointerException();
        }
        for(int i = referObjectStack.size() - 1; i >= 0; i--){
            if(referObjName.equals(referObjectStack.get(i).name)){
                return true;
            }
        }
        return false;
    }


    public void push(ContextReferObject referObject) {
        if(referObject == null){
            throw new NullPointerException();
        }
        this.referObjectStack.add(referObject);
    }

    public void pop() {
        this.referObjectStack.remove(this.referObjectStack.size() - 1);
    }


    /**
     * Following action mainly do variable validation including:
     *      check if the property name given in variable and express valid
     *      use reflection to find the getter method for the needed property and save for later use
     * @param root
     */
    public void prepareFor(DocElementUnit root) {
        // Nothing to do, use laze load mode to load getter method
    }


    public Object getPropertyValue(String referObjName, String propertyName) throws InvocationTargetException, IllegalAccessException {
        Object target = null;
        if(referObjName == null){
            target = rootObject;
        }else{
            target = this.getReferObject(referObjName);
        }

        if(target == null){
            throw new NullPointerException("Try to get property ["+propertyName+"] but no object defined with given name: " + referObjName);
        }

        Method getMethod = this.getObjectGetMethod(target.getClass(), propertyName);
        return getMethod.invoke(target);
    }

    public String getPropertyValueAsString(String referObjName, String propertyName) throws InvocationTargetException, IllegalAccessException {
        Object propertyValue = this.getPropertyValue(referObjName, propertyName);

        if(propertyValue == null){
            return "";
        }else{
            return Objects.toString(propertyValue);
        }
    }


//    public String getVariableValue(PlaceholderVariable var) throws InvocationTargetException, IllegalAccessException {
//        Object propertyValue = this.getPropertyValue(var.getReferObjName(), var.getPropertyName());
//
//        if(propertyValue == null){
//            return "";
//        }else{
//            return Objects.toString(propertyValue);
//        }
//    }

    private Method getObjectGetMethod(Class<?> aClass, String propertyName) {
        Map<String, Method> methodMap = this.getterMethodMap.get(aClass);
        if(methodMap == null){
            methodMap = new HashMap<>();
            this.getterMethodMap.put(aClass, methodMap);
        }


        Method method = methodMap.get(propertyName);
        if(method == null){
            try {
                method = findMethodFromClass(aClass, propertyName);
                methodMap.put(propertyName, method);
            } catch (NoSuchMethodException e) {
                throw new IllegalArgumentException("Not able to get the getter method for property '"
                        + propertyName + "' on class: " + aClass + ". The getter method must be public and name like 'getXXXX'" +
                        " where XXXX is capitalized property name", e);
            }
        }
        return method;
    }

    private Method findMethodFromClass(Class<?> clazz, String propertyName) throws NoSuchMethodException {

        String capitalizedName = StringUtils.capitalize(propertyName);
        String getterName = "get" + capitalizedName;
        Method getterMethod = null;
        try {
            getterMethod = clazz.getMethod(getterName);
            return getterMethod;
        } catch (NoSuchMethodException e) {
            try {
                getterName = "is" + capitalizedName;
                getterMethod = clazz.getMethod(getterName);
                return getterMethod;
            } catch (NoSuchMethodException ee) {
                throw new NoSuchMethodException("Neither get" + capitalizedName  + "() nor is" + capitalizedName + " " +
                        "found on class: " + clazz.getName());
            }
        }
    }

    public Object getReferObject(String referObjName) {
        if(referObjName == null){
            throw new NullPointerException();
        }

        if(ROOT_OBJECT_REFER_NAME.equals(referObjName)){
            return rootObject;
        }

        ContextReferObject obj = null;
        for(int i = referObjectStack.size() - 1; i >= 0; i--){
            obj = referObjectStack.get(i);
            if(referObjName.equals(obj.name)){
                return obj.obj;
            }
        }
        return null;
    }

}
