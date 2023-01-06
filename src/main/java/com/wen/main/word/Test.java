package com.wen.main.word;

import com.wen.main.word.eunm.Color;
import com.wen.main.word.paragraph.Paragraph;
import com.wen.main.word.table.WordTable;

import java.util.List;

public class Test {

    public static void main(String[] args) throws Exception {
        Word word = new Word(); /** word 对象 **/
        word.createNewWord(); /** 新建一个空白 word  **/
        Paragraph paragraph = new Paragraph();/** 插入段落  **/
        paragraph.addText("test001"); /** 段落插入 文字 **/
        paragraph.setpStyle("2"); /** 设置样式 2 标题 **/
        paragraph.setFontColor(Color.red.getCode());/**段落 字体 红色 **/
        paragraph.setBidi("0"); /** 固定参数 **/
        paragraph.getpPr().setAlgin("center"); /** 段落 样式 居中 **/
        word.append(paragraph); /** 将段落 插入 word **/
        WordTable wordTable = new WordTable(3,3); /**  创建 word 表格 3x3**/
        List<WordTable.Row> rows = wordTable.getRows();
        for(WordTable.Row row : rows){
            List<WordTable.Row.Cell> cells = row.getCells();
            for(WordTable.Row.Cell cell : cells){
                cell.addParagraph(paragraph);
            }
        }
        wordTable.mergeCell(1,1,2);
        wordTable.mergeRow(1,2,3);
        word.append(wordTable); /** 将表格插入文档 **/
        word.toWord("d:\\zs3.docx"); /** 输出word **/

    }

}
