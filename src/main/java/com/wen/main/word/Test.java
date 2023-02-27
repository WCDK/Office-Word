package com.wen.main.word;

import com.wen.main.word.eunm.Color;
import com.wen.main.word.paragraph.Paragraph;
import com.wen.main.word.table.WordTable;

import java.util.List;

public class Test {

    public static void main(String[] args) throws Exception {
        /** word 对象 **/
        Word word = new Word();
        /** 新建一个空白 word  **/
        word.createNewWord();
        /** 插入段落  **/
        Paragraph paragraph = new Paragraph();
        /** 段落插入 文字 **/
        paragraph.addText("test001");
        /** 设置样式 2 标题 **/
        paragraph.setpStyle("2");
        /**段落 字体 红色 **/
        paragraph.setFontColor(Color.red.getCode());
        /** 固定参数 **/
        paragraph.setBidi("0");
        /** 段落 样式 居中 **/
        paragraph.getpPr().setAlgin("center");
        /** 将段落 插入 word **/
        word.append(paragraph);
        /**  创建 word 表格 3x3**/
        WordTable wordTable = new WordTable(3,3);
//        List<WordTable.Row> rows = wordTable.getRows();
//        for(WordTable.Row row : rows){
//            List<WordTable.Row.Cell> cells = row.getCells();
//            for(WordTable.Row.Cell cell : cells){
//                cell.addParagraph(paragraph);
//            }
//        }
//        wordTable.mergeCell(1,1,3);
//        wordTable.mergeRow(1,2,3);
        /** 将表格插入文档 **/
        word.append(wordTable);
        /** 输出word **/
        word.toWord("d:\\zs3.docx");

    }

}
