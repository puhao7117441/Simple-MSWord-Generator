package puhao.simple.msword.generator.parse;


import java.util.regex.Pattern;

public class PlaceholderVariable {
    public static final String SUPPORTED_CHAR_DESC = "lower and upper case A to Z, 0 to 9, underscore and dot";
    public static final Pattern VARIABLE = Pattern.compile("\\$\\{([a-zA-Z_][a-zA-Z0-9_]*\\.)?([a-zA-Z_][a-zA-Z0-9_]*)\\}");

    public PlaceholderVariable(String var, int startIndex, int endIndex, ParsingContext context){
        this.variable = var;
        this.startIndex = startIndex;
        this.endIndex = endIndex;

        int indexOfDot = var.indexOf('.');
        if(indexOfDot > 0){
            this.referObjName = var.substring(2, indexOfDot).trim();
            this.propertyName = var.substring(indexOfDot + 1, var.length() - 1).trim();
            if(!context.hasReferObject(this.referObjName)){
                throw new IllegalArgumentException("No object with name '"+this.referObjName+"' found in context. Variable: " + var);
            }

        }else{
            this.referObjName = null;
            this.propertyName = var.substring(2,var.length() - 1).trim();
        }

        if(this.propertyName.equals("")){
            throw new IllegalArgumentException("Empty property name. Variable: " + var);
        }
    }
    String propertyName;
    String referObjName;
    String variable;
    int startIndex;
    int endIndex;

    /**
     *  valid range below ASCII code:
     *  32(space), 46(dot), 48(a) - 57(z), 65(0) - 90(9), 95(_), 97(A)-122(z)
     * @param c
     * @return
     */
    public static boolean isInvalidChar(char c){
        return c < 32 || c > 122
                || (c > 32 && c < 46)
                || (c == 47)
                || (c > 57 && c < 65)
                || (c > 90 && c < 95)
                || (c == 96);
    }

    public String getPropertyName() {
        return propertyName;
    }

    public String getReferObjName() {
        return referObjName;
    }

    public String getVariable() {
        return variable;
    }

    public int getStartIndex() {
        return startIndex;
    }

    public int getEndIndex() {
        return endIndex;
    }

    @Override
    public String toString() {
        return "PlaceholderVariable{" +
                "referObjName='" + referObjName + '\'' +
                ", variable='" + variable + '\'' +
                ", startIndex=" + startIndex +
                ", endIndex=" + endIndex +
                '}';
    }
}
