package puhao.simple.msword.generator.parse;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import puhao.simple.msword.generator.MSWordExporter;

import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

public class ParsedXWPFParagraph implements DocElementUnit, ExpressionContent{
    private XWPFParagraph paragraph;


    private int[] escapeCharIndexArr;
    private boolean hasEscapeChar;
    private boolean hasVariable;
    private boolean expression;
    private List<PlaceholderVariable> variables = new ArrayList<>();
    private ForLoopExpression forLoopExpression;
    private String text;

    private DocElementUnit nextSlibing;
    private DocElementUnit parent;

    public ParsedXWPFParagraph(XWPFParagraph paragraph, ParsingContext context, DocElementUnit parent){
        this.paragraph = paragraph;

        this.parent = parent;

        // footnote text are ignored in this solution
        this.text = paragraph.getParagraphText();

        LinkedList<Integer> escapeCharIndex = new LinkedList<>();

        char c = ' ';
        int variableStartIndex = -1, variableEndIndex = -1;
        StringBuilder sb = new StringBuilder();
        int expressStartIndex, expressEndIndex;
        PlaceholderVariable variable;
        for(int i = 0; i < text.length(); i++){

            c = text.charAt(i);
            if(c > 127){
                // here only check for escape char and dollar char, all other are skipped
                continue;
            }

            if(c == MSWordExporter.ESCAPE){
                /**
                 * Here validation is very important, must ensure only below given two case are accepted, otherwish
                 * the export phase later will fail
                 */
                if(i < text.length() - 1){
                     if(text.charAt(i + 1) == MSWordExporter.DOLLAR || text.charAt(i + 1) == MSWordExporter.ESCAPE){
                         escapeCharIndex.add(i);
                         i++;
                         this.hasEscapeChar = true;
                         continue;
                     }else{
                         throw new IllegalArgumentException("Invalid escape character, only $ and \\ need escape." +
                                 " Means you can only escape '$' by write '\\$'" +
                                 " or escape '\\' by '\\\\' in the doc. Invalid character: " + text.charAt(i + 1) +
                                 " at index " + (i + 1) + " of paragraph text: " + text);
                     }
                }else{
                    throw new IllegalArgumentException("invalid escape character ");
                }
            }else if(c == MSWordExporter.DOLLAR){
                if((i < text.length() - 1) && (text.charAt(i + 1) == MSWordExporter.BRACES_L)){
                    if(i < text.length() - 2 && text.charAt(i + 2) == MSWordExporter.BRACES_L){
                        this.expression = true;
                        if(this.forLoopExpression != null){
                            throw new IllegalArgumentException("One paragraph can maximum contain one expression, but got more than one in paragraph: " + text);
                        }
                        variableStartIndex = i;
                        sb.setLength(0);
                        sb.append(MSWordExporter.EXPRESSION_START);
                        i = extractExpression(text, i + 3, sb);
                        this.forLoopExpression = new ForLoopExpression(sb.toString(), variableStartIndex, i + 1, context);
                    }else{
                        variableStartIndex = i;
                        sb.setLength(0);
                        sb.append(MSWordExporter.VARIABLE_START);
                        try {
                            i = extractVariable(text, i + 2, sb);
                            variable = new PlaceholderVariable(sb.toString(), variableStartIndex, i + 1, context);
                            this.variables.add(variable);
                            this.hasVariable = true;
                        }catch (IllegalArgumentException e){
                            if(ForLoopExpression.ERROR_EXPRESSION_1.matcher(text).find()){
                                throw new IllegalArgumentException("Invalid variable found: [" + text
                                        + "]. It looks like contain a for-loop-expression, but for-loop-expression " +
                                        "must start with ${{ and end with }} (double braces), like ${{for item of itemList}}", e);
                            }else{
                                throw e;
                            }
                        }
                    }
                }
            }
        }

        if(this.expression && this.hasVariable){
            throw new IllegalArgumentException("One paragraph can not contain both expression and variable. expression should in its own paragraph. Paragraph text: " + text);
        }
        if(this.hasEscapeChar){
            this.escapeCharIndexArr = escapeCharIndex.stream().mapToInt(i -> i.intValue()).toArray();
        }
    }

    /**
     * The start index i is first character after {.
     * If the variable is ${abc}, then the index i is point to char 'a' in this string.
     * @param text
     * @param i
     * @param sb
     * @return
     */
    private static int extractVariable(String text, int i, StringBuilder sb) {
        char c = ' ';
        boolean dotFound = false;
        for(; i < text.length(); i++){
            c = text.charAt(i);

            if(c == ' ' || c == '\t'){
                throw new IllegalArgumentException("Variable name must not contain space character nor tab character. " +
                        "The supported char for variable name are: "+PlaceholderVariable.SUPPORTED_CHAR_DESC+". Paragraph: " + text);
            }

            if(c == '.'){
                if(dotFound){
                    throw new IllegalArgumentException("Invalid variable, for one variable, only one dot character can be used to identify the object. Paragraph: " + text);
                }else{
                    dotFound = true;
                }
            }

            sb.append(c);

            if(c == MSWordExporter.BRACES_R){
                if(sb.length() == 0){
                    throw new IllegalArgumentException("Variable must have a name, like ${age}. The supported char for variable name are: "
                            +PlaceholderVariable.SUPPORTED_CHAR_DESC+". Paragraph: " + text);
                }
                return i;
            }else if(PlaceholderVariable.isInvalidChar(c)){
                throw new IllegalArgumentException("Variable name contain invalid character (only support "+
                        PlaceholderVariable.SUPPORTED_CHAR_DESC+") or variable missing right braces at index "+i+" of paragraph: " + text);
            }
        }
        throw new IllegalArgumentException("Variable missing right braces of paragraph: " + text);
    }

    private int extractExpression(String text, int i, StringBuilder sb) {
        char c = ' ';
        boolean empty = true;
        for(; i < text.length(); i++){
            c = text.charAt(i);

            sb.append(c);
            // valid range
            // 32, 46, 48 - 57, 65 - 90, 95, 97-122

            if(c == MSWordExporter.BRACES_R){
                if(i < text.length() - 1 && text.charAt(i + 1) == MSWordExporter.BRACES_R){
                    if(empty){
                        throw new IllegalArgumentException("Expression must not empty. Expression should like ${{let p of persons}}. Paragraph: " + text);
                    }
                    sb.append(MSWordExporter.BRACES_R);
                    return i + 1;
                }else{
                    throw new IllegalArgumentException("Unexpected right braces. Expression must end with double right braces, like ${{let p of persons}}. Paragraph: " + text);
                }
            }else if(c < 32 || c > 122
                    || (c > 32 && c < 46)
                    || (c == 47)
                    || (c > 57 && c < 65)
                    || (c > 90 && c < 95)
                    || (c == 96)){
                throw new IllegalArgumentException("Expression contain invalid character (only support space, lower and upper case A to Z, 0 to 9, underscore and dot) or variable missing right braces at index "+i+" of paragraph: " + text);
            }
            empty = false;
        }
        throw new IllegalArgumentException("Expression missing double right braces of paragraph: " + text);
    }



    public XWPFParagraph getParagraph() {
        return paragraph;
    }



    public void setParagraph(XWPFParagraph paragraph) {
        this.paragraph = paragraph;
    }

    public int[] getEscapeCharIndexArr() {
        return escapeCharIndexArr;
    }

    public void setEscapeCharIndexArr(int[] escapeCharIndexArr) {
        this.escapeCharIndexArr = escapeCharIndexArr;
    }

    public boolean isHasEscapeChar() {
        return hasEscapeChar;
    }

    public void setHasEscapeChar(boolean hasEscapeChar) {
        this.hasEscapeChar = hasEscapeChar;
    }

    public boolean isHasVariable() {
        return hasVariable;
    }

    public void setHasVariable(boolean hasVariable) {
        this.hasVariable = hasVariable;
    }

    public boolean isExpression() {
        return expression;
    }

    public void setExpression(boolean expression) {
        this.expression = expression;
    }

    public List<PlaceholderVariable> getVariables() {
        return variables;
    }

    public void setVariables(List<PlaceholderVariable> variables) {
        this.variables = variables;
    }

    public ForLoopExpression getForLoopExpression() {
        return forLoopExpression;
    }

    public void setForLoopExpression(ForLoopExpression forLoopExpression) {
        this.forLoopExpression = forLoopExpression;
    }

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

    public void setNextSibling(DocElementUnit unit) {
        this.nextSlibing = unit;
    }

    @Override
    public String toString() {
        return "ParsedXWPFParagraph{" +
                "hasEscapeChar=" + hasEscapeChar +
                ", hasVariable=" + hasVariable +
                ", hasExpression=" + expression +
                ", variables=" + variables +
                ", expression=" + forLoopExpression +
                ", text='" + text + '\'' +
                '}';
    }


    @Override
    public DocElementType getType() {
        return DocElementType.PARAGRAPH;
    }

    @Override
    public Object getRawDocObject() {
        return this.getParagraph();
    }

    @Override
    public DocElementUnit parent() {
        return parent;
    }

    @Override
    public boolean hasNextSibling() {
        return nextSlibing != null;
    }

    @Override
    public DocElementUnit nextSibling() {
        return nextSlibing;
    }

    @Override
    public boolean isContainerUnit() {
        return false;
    }

    @Override
    public List<DocElementUnit> getChildUnit() {
        return null;
    }
}
