package puhao.simple.msword.generator.parse;

import puhao.simple.msword.generator.MSWordHandlerContext;

import java.util.ArrayList;
import java.util.List;

public class ParsingContext {
    private int expressStartCountCheck = 0;
    private List<String> referObjectNameStack = new ArrayList<>();

    public ParsingContext(){
    }

    public void pushNewReferObjectName(String referObjName){
        if(referObjName == null){
            throw new NullPointerException();
        }
        this.referObjectNameStack.add(referObjName);
    }

    public void popReferObjectName(){
        this.referObjectNameStack.remove(referObjectNameStack.size() - 1);
    }

    public boolean hasReferObject(String referObjName) {
        if(referObjName == null){
            throw new NullPointerException();
        }

        if(MSWordHandlerContext.ROOT_OBJECT_REFER_NAME.equals(referObjName)){
            return true;
        }

        for(int i = referObjectNameStack.size() - 1; i >= 0; i--){
            if(referObjName.equals(referObjectNameStack.get(i))){
                return true;
            }
        }
        return false;
    }

    public boolean isExpressCountBalanced(){
        return this.expressStartCountCheck == 0;
    }
    public void countExpressStart(){
        this.expressStartCountCheck++;
    }
    public void countExpressEnd(){
        this.expressStartCountCheck--;
    }
}
