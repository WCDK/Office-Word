package com.wcdk.main.word;

import com.wcdk.main.word.image.WordImage;
import com.wcdk.main.word.paragraph.Paragraph;
import com.wcdk.main.word.core.eunm.Algin;
import com.wcdk.main.word.core.eunm.Color;

public class Test {

    public static void main(String[] args) throws Exception {
//        int i = MyUnits.pxToEMU(4096);
//        System.out.println(i);
        a();
//        b();
    }
    public static void b() throws Exception{
        Word word = new Word();
        word.loadWord("d:\\zs4.docx");
        word.toWord("d:\\zs5.docx");
        System.out.println(word.countWords());
    }
    public static void a() throws Exception{
        /** word 对象 **/
        Word word = new Word();
        /** 新建一个空白 word  **/
        word.createNewWord();
        /** 插入段落  **/
        Paragraph paragraph = new Paragraph();
        /** 段落插入 文字 **/
        paragraph.setText("test001");
        /** 设置样式 2 标题 **/
        paragraph.setStyle("2");
        /**段落 字体 红色 **/
        paragraph.setFontColor(Color.red.getCode());
        /** 段落 样式 居中 **/
        paragraph.setAlgin(Algin.center);
        /** 将段落 插入 word **/
        word.append(paragraph);

        WordImage wordImage = new WordImage("E:\\2\\3.jpg");
        wordImage.setAlgin(Algin.center);
//        WordImage wordImage = new WordImage("e:/234.jpg");
        word.append(wordImage);
        /** 输出word **/
        word.toWord("d:\\zs4.docx");
    }

}
