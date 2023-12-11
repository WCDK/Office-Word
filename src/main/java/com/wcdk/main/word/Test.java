package com.wcdk.main.word;

import com.wcdk.main.word.image.WordImage;
import com.wcdk.main.word.paragraph.Paragraph;
import com.wcdk.main.core.eunm.Algin;
import com.wcdk.main.core.eunm.Color;

public class Test {

    public static void main(String[] args) throws Exception {
        a();
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
        word.createNewWord(true,"1");
        /** 插入段落  **/
        Paragraph paragraph = new Paragraph();
        /** 段落插入 文字 **/
        paragraph.setText("test001");
        /** 设置样式  1 正文 大于等于2标题  2:一级标题 3 二级标题 以此类推**/
        paragraph.setStyle("2");
        /**段落 字体 红色 **/
        paragraph.setFontColor(Color.red.getCode());
        /** 段落 样式 居中 **/
        paragraph.setAlgin(Algin.center);
        paragraph.setBidi("0");
        Paragraph paragraph1 = new Paragraph();

        /** 段落插入 文字 **/
        paragraph1.setText("test001");
        /** 设置样式  1 正文 大于等于2标题  2:一级标题 3 二级标题 以此类推**/
        paragraph1.setStyle("3");
        /**段落 字体 红色 **/
        paragraph1.setFontColor(Color.red.getCode());
        /** 段落 样式 居中 **/
        paragraph1.setAlgin(Algin.center);
        paragraph1.setBidi("0");

//        WordImage wordImage = new WordImage("E:\\2\\3.jpg");
//        wordImage.setAlgin(Algin.center);
////        WordImage wordImage = new WordImage("e:/234.jpg");
//        word.append(wordImage);

        word.append(paragraph,paragraph1);
        /** 输出word **/
        word.toWord("d:\\zs5.docx");
    }

}
