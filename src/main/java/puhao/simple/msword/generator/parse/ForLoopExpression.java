package puhao.simple.msword.generator.parse;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * A for loop expression like:
 *   ${{for loopItemName of iterable}}
 *
 * A for loop expression must start with '${{' and end with '}}'
 * inside the braces two mandatory key words are: 'for' and 'of'.
 * the loopItemName is must follow regex: [a-zA-Z_][a-zA-Z0-9_]*
 * the iterable object name must follow regex:
 *
 *
 */
public class ForLoopExpression {
    public static final Pattern EXPRESSION_END = Pattern.compile("\\$\\{\\{\\s*end\\s*\\}\\}");
    public static final Pattern EXPRESSION_DEFINE = Pattern.compile("\\$\\{\\{\\s*for\\s+([a-zA-Z][a-zA-Z0-9_]*)\\s+of\\s+([a-zA-Z_][a-zA-Z0-9_]*\\.)?([a-zA-Z_][a-zA-Z0-9_]*)\\s*\\}\\}");

    /**
     * This ERROR_EXPRESSION_1 is used to enhance error message.
     * some time people may mission one brace for the expression, like
     * the correct:        ${{for loopItemName of iterable}}
     * missing on brace:  ${for loopItemName of iterable}
     *
     * In this case, use this ERROR_EXPRESSION_1 pattern can detect this error
     * and provide user friendly error message
     */
    public static final Pattern ERROR_EXPRESSION_1 = Pattern.compile("\\$\\{\\s*for\\s+([a-zA-Z][a-zA-Z0-9_]*)\\s+of\\s+([a-zA-Z_][a-zA-Z0-9_]*\\.)?([a-zA-Z_][a-zA-Z0-9_]*)\\s*\\}");

    public static final int EXPRESSION_LOOP_ITEM_GROUP_ID = 1;
    public static final int EXPRESSION_ITERABLE_OBJ_PARENT_GROUP_ID = 2;
    public static final int EXPRESSION_ITERABLE_OBJ_GROUP_ID = 3;
    String expression;
    String iterableObjReferObjectName;
    String iterablePropertyName;
    String loopItemName;
    int startIndex;
    int endIndex;
    boolean isStart;
    boolean isEnd;
    public ForLoopExpression(String expression, int startIndex, int endIndex, ParsingContext context) {
        this.expression = expression;
        this.startIndex = startIndex;
        this.endIndex = endIndex;


        if(EXPRESSION_END.matcher(expression).matches()){
            isStart = false;
            isEnd = true;
            context.countExpressEnd();
        }else{
            Matcher matcher = EXPRESSION_DEFINE.matcher(expression);
            if(matcher.matches()){
                /**
                 * Example:
                 * ${{for t of ADFAF.trainings}} =>
                 *  group 1 = t
                 *  group 2 = ADFAF.
                 *  group 3 = trainings
                 */
                loopItemName = matcher.group(EXPRESSION_LOOP_ITEM_GROUP_ID);
                iterableObjReferObjectName = matcher.group(EXPRESSION_ITERABLE_OBJ_PARENT_GROUP_ID);
                iterablePropertyName = matcher.group(EXPRESSION_ITERABLE_OBJ_GROUP_ID);

                if(iterableObjReferObjectName != null){
                    // remove the last dot char
                    iterableObjReferObjectName = iterableObjReferObjectName.substring(0, iterableObjReferObjectName.length() - 1);
                }

                if(iterableObjReferObjectName != null && !context.hasReferObject(iterableObjReferObjectName)){
                    throw new IllegalArgumentException("No object with name '"+ iterableObjReferObjectName +"' found in context. For-loop-expression: " + expression);
                }
                context.pushNewReferObjectName(loopItemName);

                isStart = true;
                isEnd = false;
                context.countExpressStart();
            }else{
                throw new IllegalArgumentException("Invalid for-loop-expression: " + expression);
            }

        }

    }

    public String getExpression() {
        return expression;
    }

    public void setExpression(String expression) {
        this.expression = expression;
    }

    public String getIterableObjReferObjectName() {
        return iterableObjReferObjectName;
    }

    public void setIterableObjReferObjectName(String iterableObjReferObjectName) {
        this.iterableObjReferObjectName = iterableObjReferObjectName;
    }

    public String getIterablePropertyName() {
        return iterablePropertyName;
    }

    public void setIterablePropertyName(String iterablePropertyName) {
        this.iterablePropertyName = iterablePropertyName;
    }

    public String getLoopItemName() {
        return loopItemName;
    }

    public void setLoopItemName(String loopItemName) {
        this.loopItemName = loopItemName;
    }

    public int getStartIndex() {
        return startIndex;
    }

    public void setStartIndex(int startIndex) {
        this.startIndex = startIndex;
    }

    public int getEndIndex() {
        return endIndex;
    }

    public void setEndIndex(int endIndex) {
        this.endIndex = endIndex;
    }

    public boolean isStart() {
        return isStart;
    }

    public void setStart(boolean start) {
        isStart = start;
    }

    public boolean isEnd() {
        return isEnd;
    }

    public void setEnd(boolean end) {
        isEnd = end;
    }

    @Override
    public String toString() {
        return "ForLoopExpression{" +
                "expression='" + expression + '\'' +
                ", iterableObjParentName='" + iterableObjReferObjectName + '\'' +
                ", iterableObjName='" + iterablePropertyName + '\'' +
                ", loopItemName='" + loopItemName + '\'' +
                ", startIndex=" + startIndex +
                ", endIndex=" + endIndex +
                ", isStart=" + isStart +
                ", isEnd=" + isEnd +
                '}';
    }
}
