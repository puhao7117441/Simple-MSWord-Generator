package puhao.simple.msword.generator;

import org.apache.poi.xwpf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import puhao.simple.msword.generator.parse.*;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Pattern;

/**
 * Once the instance be created and template have no change, this instance can be reuse.
 * This class is NOT thread-safe, don't use it multi-threads env.
 */
public class MSWordExporter<T> {
    public static final DateTimeFormatter DOC_CREATION_TIME_FORMATTER = DateTimeFormatter.ofPattern("uuuu年MM月dd日 HH:mm:ss");
    private static final Logger logger = LoggerFactory.getLogger("MSWordExporter");
    private static final Pattern TEMPLATE_VARIABLE_PATTERN = Pattern.compile("\\$\\{([a-zA-Z_][a-zA-Z0-9_]*\\.)?([a-zA-Z_][a-zA-Z0-9_]*)\\}");
    public static final String VARIABLE_START = "${";
    public static final String EXPRESSION_START = "${{";
    public static final char DOLLAR = '$';
    public static final char ESCAPE = '\\';
    public static final char BRACES_L = '{';
    public static final char BRACES_R = '}';
    public static final char DELETE = 127;

    private Path templatePath;
    private ParsedXWPFDocument parsedDocument;

    public MSWordExporter(Path templatePath) throws IOException {
        this.templatePath = templatePath;
        try(XWPFDocument doc = new XWPFDocument(Files.newInputStream(templatePath, StandardOpenOption.READ))) {
            this.parsedDocument = new ParsedXWPFDocument(doc);
        }
    }



    public void exportTo(Object rootObject, Path targetFilePath) throws IOException, InvocationTargetException, IllegalAccessException {
        MSWordHandlerContext context = new MSWordHandlerContext(rootObject);
        context.prepareFor(parsedDocument);

        try(XWPFDocument newDoc = createNewEmptyDoc()){
            headerAndFooterHandler(newDoc, parsedDocument, context);

            List<? extends DocElementUnit> elementUnits = parsedDocument.getChildUnit();
            for(int i = 0; i < elementUnits.size(); ){
                int parsedSiblingCount = copyAndFill(elementUnits.get(i), context, newDoc);
                i = i + parsedSiblingCount;
            }

            this.writeDoc(newDoc, targetFilePath);
        }

    }

    private void headerAndFooterHandler(XWPFDocument newDoc, ParsedXWPFDocument parsedDoc, MSWordHandlerContext context) throws InvocationTargetException, IllegalAccessException {
        {
            List<ParsedXWPFHeader> parsedHeaderList = parsedDoc.getParsedXWPFHeaders();
            List<XWPFHeader> headerList = newDoc.getHeaderList();
            for (int i = 0; i < parsedHeaderList.size(); i++) {
                ParsedXWPFHeader parsedHeader = parsedHeaderList.get(i);
                for (int pi = 0; pi < parsedHeader.getParagraphs().size(); pi++) {
                    this.simpleParagraphFill(parsedHeader.getParagraph(pi)
                            , headerList.get(i).getParagraphArray(pi)
                            , context);
                }
            }
        }
        {
            List<ParsedXWPFFooter> parsedFooterList = parsedDoc.getParsedXWPFFooters();
            List<XWPFFooter> footerList = newDoc.getFooterList();
            for (int i = 0; i < parsedFooterList.size(); i++) {
                ParsedXWPFFooter parsedFooter = parsedFooterList.get(i);
                for (int pi = 0; pi < parsedFooter.getParagraphs().size(); pi++) {
                    this.simpleParagraphFill(parsedFooter.getParagraph(pi)
                            , footerList.get(i).getParagraphArray(pi)
                            , context);
                }
            }
        }
    }

    private XWPFDocument createNewEmptyDoc() throws IOException {
        /**
         * Since create a new document then copy document style from template to the
         * new doc is very complex (anyway it's very complex for me), I use different
         * solution: Read the template file again, and delete all bodies in that element
         * (except header and footer).
         */
        XWPFDocument doc = new XWPFDocument(Files.newInputStream(templatePath, StandardOpenOption.READ));
        while(doc.getBodyElements().size() > 0){
            doc.removeBodyElement(0);
        }
        //POIHelper.cloneDocumentStyles(newDoc, (XWPFDocument) root.getRawDocObject());
        return doc;
    }


    private int copyAndFill(DocElementUnit element
            , MSWordHandlerContext context
            , XWPFDocument newDoc) throws InvocationTargetException, IllegalAccessException {
        int parsedSiblingCount = 0;
        if(element.getType() == DocElementType.PARAGRAPH){
            ParsedXWPFParagraph parsedParagraph = (ParsedXWPFParagraph) element;
            if(parsedParagraph.isExpression()){
                if(parsedParagraph.getForLoopExpression().isStart()) {
                    parsedSiblingCount = handleNormalExpressionStart(newDoc, parsedParagraph, context);
                }else{
                    throw new IllegalStateException("No chance to get here, the end expression should already skipped");
                }
            }else{
                XWPFParagraph newP  = newDoc.createParagraph();
                POIHelper.cloneParagraph(newP, parsedParagraph.getParagraph());

                this.simpleParagraphFill(parsedParagraph, newP, context);

                parsedSiblingCount++;
            }
        }else if(element.getType() == DocElementType.TABLE){

            handleTable(newDoc, (ParsedTable) element, context);

            parsedSiblingCount++;
        }else if(element.getType() == DocElementType.TABLE_ROW){
            throw new UnsupportedOperationException("Shouldn't happen");
        }else if(element.getType() == DocElementType.TABLE_CELL){
            throw new UnsupportedOperationException("Shouldn't happen");
        }else if(element.getType() == DocElementType.OTHER){
            throw new UnsupportedOperationException("Shouldn't happen");
        }else{
            throw new IllegalArgumentException("Unknown element type: " + element.getType());
        }


        return parsedSiblingCount;
    }

    private void simpleParagraphFill(ParsedXWPFParagraph parsedParagraph
            , XWPFParagraph newP
            , MSWordHandlerContext context)
            throws InvocationTargetException, IllegalAccessException {
        if(parsedParagraph.isHasVariable()){
            this.fillVariableValueToParagraph(newP, parsedParagraph , context);
        }else if(parsedParagraph.isHasEscapeChar()){
            this.removeEscapeCharInParagraph(newP, parsedParagraph.getEscapeCharIndexArr());
        }
    }

    private void handleTable(XWPFDocument newDoc, ParsedTable parsedTable, MSWordHandlerContext context) throws InvocationTargetException, IllegalAccessException {
        XWPFTable newTable = newDoc.createTable();

        if(parsedTable.isHasExpression()){
            expressionTableHandler(newTable, parsedTable, newDoc, context);
            //newly created table has one row by default. we need to remove the default row.
            newTable.removeRow(0);
        }else{
            //This cloneTable will also remove the default row 0
            POIHelper.cloneTable(newTable, (XWPFTable) parsedTable.getRawDocObject());
            this.simpleTableFill(newTable, parsedTable, context, newDoc);
        }
    }

    private void expressionTableHandler(XWPFTable newTable, ParsedTable parsedTable
            , XWPFDocument newDoc
            , MSWordHandlerContext context) throws InvocationTargetException, IllegalAccessException {
        for(int rowIndex = 0; rowIndex < parsedTable.getRows().size();){
            ParsedTableRow parsedRow = parsedTable.getRow(rowIndex);
            rowIndex += expressionTableRowHandler(parsedRow, newTable, context);
        }
    }

    private int expressionTableRowHandler(ParsedTableRow parsedRow
            , XWPFTable newTable
            , MSWordHandlerContext context) throws InvocationTargetException, IllegalAccessException {

        if(parsedRow.isExpression()){
            ForLoopExpression expression = parsedRow.getForLoopExpression();
            if(expression.isEnd()){
                throw new IllegalStateException("No chance to get here, the end expression should already skipped");
            }
            return handlerExpressionStartTableRow(newTable, parsedRow,context);
        }else{
            XWPFTableRow newRow = newTable.createRow();
            POIHelper.cloneRow(newRow, (XWPFTableRow) parsedRow.getRawDocObject());
            this.simpleRowFill(parsedRow, newRow, context);
            return 1;
        }
    }


    private int handlerExpressionStartTableRow(XWPFTable newTable, ParsedTableRow startRow
            , MSWordHandlerContext context) throws InvocationTargetException, IllegalAccessException {
        ForLoopExpression expression = startRow.getForLoopExpression();

        Object iterableObj = context.getPropertyValue(expression.getIterableObjReferObjectName()
                , expression.getIterablePropertyName());

        if(iterableObj == null){
            throw new NullPointerException("For-loop-expression on null. '" +expression.getIterablePropertyName()+"' is null. Expression: " + expression.getExpression());
        }
        if(!(iterableObj instanceof Iterable)){
            throw new IllegalArgumentException("For-loop-expression on a non-iterable object but '" + iterableObj.getClass() + "'. Expression: " + expression.getExpression());
        }

        ContextReferObject referObj = new ContextReferObject(expression.getLoopItemName(), null);
        context.push(referObj);

        Iterable iterableVal = (Iterable) iterableObj;
        int totalIterableItemCount = 0;
        for(Object obj : iterableVal){totalIterableItemCount++;}

        List<DocElementUnit> allExpressContentList = findExpressionContent(startRow);

        if(totalIterableItemCount == 0){
            return allExpressContentList.size();
        }else{
            for(Object obj : iterableVal){
                referObj.obj = obj;
                /**
                 * i = 0 is the expression start
                 * i = allExpressContentList.size() - 1 is the expression end
                 * to do copy and fill we need ignore the start and end
                 */
                int innerSiblingCount = 0;
                for(int i = 1; i < allExpressContentList.size() - 1;){
                    innerSiblingCount = expressionTableRowHandler((ParsedTableRow) allExpressContentList.get(i)
                            , newTable, context);
                    i += innerSiblingCount;
                }
            }
        }

        return allExpressContentList.size();
    }

    private void simpleTableFill(XWPFTable newTable, ParsedTable parsedTable
            , MSWordHandlerContext context, XWPFDocument newDoc) throws InvocationTargetException, IllegalAccessException {
        for(int rowIndex = 0; rowIndex < parsedTable.getRows().size(); rowIndex++){
            ParsedTableRow parsedRow = parsedTable.getRow(rowIndex);
            XWPFTableRow row = newTable.getRow(rowIndex);
            this.simpleRowFill(parsedRow, row, context);
        }
    }

    private void simpleRowFill(ParsedTableRow parsedRow, XWPFTableRow row, MSWordHandlerContext context) throws InvocationTargetException, IllegalAccessException {
        for(int cellIndex = 0;cellIndex < parsedRow.getCells().size(); cellIndex++){
            ParsedTableCell parsedCell = parsedRow.getCells().get(cellIndex);
            XWPFTableCell cell = row.getCell(cellIndex);
            this.simpleCellFill(parsedCell, cell, context);
        }
    }

    private void simpleCellFill(ParsedTableCell parsedCell, XWPFTableCell cell, MSWordHandlerContext context) throws InvocationTargetException, IllegalAccessException {
        for(int pIndex = 0; pIndex < parsedCell.getParagraphs().size(); pIndex++) {
            this.simpleParagraphFill(parsedCell.getParagraphs(pIndex)
                    , cell.getParagraphArray(pIndex)
                    , context);
        }
    }


    private int handleNormalExpressionStart(XWPFDocument newDoc, ParsedXWPFParagraph expressionStart
            , MSWordHandlerContext context) throws InvocationTargetException, IllegalAccessException {
        ForLoopExpression expression = expressionStart.getForLoopExpression();

        Object iterableObj = context.getPropertyValue(expression.getIterableObjReferObjectName()
                                                    , expression.getIterablePropertyName());

        if(iterableObj == null){
            throw new NullPointerException("For-loop-expression on null. '" +expression.getIterablePropertyName()+"' is null. Expression: " + expression.getExpression());
        }
        if(!(iterableObj instanceof Iterable)){
            throw new IllegalArgumentException("For-loop-expression on a non-iterable object but '" + iterableObj.getClass() + "'. Expression: " + expression.getExpression());
        }

        ContextReferObject referObj = new ContextReferObject(expression.getLoopItemName(), null);
        context.push(referObj);

        Iterable iterableVal = (Iterable) iterableObj;
        int totalIterableItemCount = 0;
        for(Object obj : iterableVal){totalIterableItemCount++;}

        List<DocElementUnit> allExpressContentList = findExpressionContent(expressionStart);

        if(totalIterableItemCount == 0){
            return allExpressContentList.size();
        }else{
            for(Object obj : iterableVal){
                referObj.obj = obj;
                /**
                 * i = 0 is the expression start
                 * i = allExpressContentList.size() - 1 is the expression end
                 * to do copy and fill we need ignore the start and end
                 */
                int innerSiblingCount = 0;
                for(int i = 1; i < allExpressContentList.size() - 1;){
                    innerSiblingCount = copyAndFill(allExpressContentList.get(i), context, newDoc);
                    i += innerSiblingCount;
                }
            }
        }

        return allExpressContentList.size();
    }

    private List<DocElementUnit> findExpressionContent(DocElementUnit expressionStart) {
        List<DocElementUnit> expressionContent = new ArrayList<>();
        DocElementUnit sibling = expressionStart;
        expressionContent.add(sibling);

        Stack<DocElementUnit> expressStack = new Stack<>();
        while(null != (sibling = sibling.nextSibling())){
            expressionContent.add(sibling);

            if(sibling.getType() == DocElementType.PARAGRAPH || sibling.getType() == DocElementType.TABLE_ROW){
                ExpressionContent expressionItem = (ExpressionContent)sibling;

                if(expressionItem.isExpression()){
                    if(expressionItem.getForLoopExpression().isStart()){
                        expressStack.push(sibling);
                    }else{
                        if(expressStack.size() == 0) {
                            return expressionContent;
                        }else{
                            expressStack.pop();
                        }
                    }
                }
            }
        }

        String expression = null;
        if(expressionStart instanceof ParsedXWPFParagraph){
            expression = ((ParsedXWPFParagraph)expressionStart).getForLoopExpression().getExpression();
        }else if(expressionStart instanceof ParsedTableRow){
            expression = ((ParsedTableRow)expressionStart).getForLoopExpression().getExpression();
        }else{
            expression = "<unknown>";
        }
        throw new IllegalArgumentException("Not able to find expression end of expression: "
                + expression);
    }

    private void fillVariableValueToParagraph(XWPFParagraph paragraph
            , ParsedXWPFParagraph parsedParagraph
            , MSWordHandlerContext context) throws InvocationTargetException, IllegalAccessException {
        String text = null;

        if(parsedParagraph.isHasEscapeChar()){
            text = this.replaceEscapeCharToDelete(paragraph.getParagraphText(), parsedParagraph.getEscapeCharIndexArr());
        }else{
            text = paragraph.getParagraphText();
        }

        List<PlaceholderVariable> variables = parsedParagraph.getVariables();
        List<String> varValues = new ArrayList<>(variables.size());

        String varValue = null;
        for(PlaceholderVariable v : variables){
            /**
             * If an variable not given a refer object, it have two possibility:
             * 1. The given property name is root object property
             * 2. The given property name actually is a object refer name. Here need invoke toString() of that object
             */
            if(v.getReferObjName() == null){
                Object referObj = context.getReferObject(v.getPropertyName());
                varValues.add(referObj == null?
                        context.getPropertyValueAsString(null, v.getPropertyName()): Objects.toString(referObj));
            }else{
                varValues.add(context.getPropertyValueAsString(v.getReferObjName(), v.getPropertyName()));
            }
        }

        String replacedText = replaceVariableInText(text, parsedParagraph.getVariables(), varValues);

        if(parsedParagraph.isHasEscapeChar()){
            StringBuilder sb = new StringBuilder();
            char c = ' ';
            for(int i = 0; i < replacedText.length(); i++){
                c = replacedText.charAt(i);
                if(c == DELETE){
                    continue;
                }
                sb.append(c);
            }

            replacedText = sb.toString();
        }

        setParagraphText(paragraph, replacedText);
    }


    private String replaceVariableInText(String text, List<PlaceholderVariable> variables
            , List<String> values){
        StringBuilder sb = new StringBuilder(text);

        int offset = 0;
        int valueLength = 0;
        int variableLength = 0;
        String value = "";
        PlaceholderVariable var = null;
        for(int i = 0; i < variables.size(); i++){
            var = variables.get(i);
            value = values.get(i);

            variableLength = var.getVariable().length();
            valueLength = value.length();


            sb.replace(var.getStartIndex() + offset,var.getEndIndex() + offset, value);

            offset += (valueLength - variableLength);
        }

        return sb.toString();
    }




    private void removeEscapeCharInParagraph(XWPFParagraph paragraph, int[] escapeCharIndexArr) {
        String text = paragraph.getParagraphText();
        StringBuilder sb = new StringBuilder(text);
        /**
         * Here the logic is highly dependent on the parse phase check, all the invalid use cases
         * of escape character already excluded (by throw exception) in ParsedXWPFParagraph
         */
        int offset = 0;
        for(int index : escapeCharIndexArr){
            sb.deleteCharAt(index + offset);
            offset--;
        }

        setParagraphText(paragraph, sb.toString());
    }

    private String replaceEscapeCharToDelete(String text, int[] escapeCharIndexArr) {
        StringBuilder sb = new StringBuilder(text);
        /**
         * Here the logic is highly dependent on the parse phase check, all the invalid use cases
         * of escape character already excluded (by throw exception) in ParsedXWPFParagraph
         */
        for(int index : escapeCharIndexArr){
            sb.setCharAt(index, DELETE);
        }

        return sb.toString();
    }

    private static void setParagraphText(XWPFParagraph paragraph, String text) {
        if(paragraph.getRuns() == null || paragraph.getRuns().size() == 0){
            paragraph.createRun();
        }else {
            while (paragraph.getRuns().size() > 1) {
                paragraph.removeRun(1);
            }
        }

        XWPFRun run = paragraph.getRuns().get(0);

        /**
         * XWPFRun.setText(String value) actually means append text to the end
         * must give the pos number to replace all text
         */
        run.setText(text == null?"":text, 0);
    }

    private void writeDoc(XWPFDocument doc, Path outputFilePath) throws IOException {
        LocalDateTime nowTime = LocalDateTime.now();

        try(OutputStream fileOutputStream = Files.newOutputStream(outputFilePath, StandardOpenOption.CREATE_NEW)) {
            doc.write(fileOutputStream);
        }
    }
}
