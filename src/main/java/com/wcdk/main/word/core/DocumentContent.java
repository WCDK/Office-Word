package com.wcdk.main.word.core;

import com.wcdk.main.word.core.eunm.PaperType;
import com.wcdk.main.word.image.WordImage;
import lombok.Data;
import org.dom4j.Element;

import java.util.ArrayList;
import java.util.List;

@Data
public class DocumentContent {
    String wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
    String cx="http://schemas.microsoft.com/office/drawing/2014/chartex";
    String cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex";
    String cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex";
    String cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex";
    String cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex";
    String cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex";
    String cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex";
    String cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex";
    String cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex";
    String mc="http://schemas.openxmlformats.org/markup-compatibility/2006";
    String aink="http://schemas.microsoft.com/office/drawing/2016/ink";
    String am3d="http://schemas.microsoft.com/office/drawing/2017/model3d";
    String o="urn:schemas-microsoft-com:office:office";
    String oel="http://schemas.microsoft.com/office/2019/extlst";
    String r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    String m="http://schemas.openxmlformats.org/officeDocument/2006/math";
    String v="urn:schemas-microsoft-com:vml";
    String wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
    String wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    String w10="urn:schemas-microsoft-com:office:word";
    String w="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    String w14="http://schemas.microsoft.com/office/word/2010/wordml";
    String w15="http://schemas.microsoft.com/office/word/2012/wordml";
    String w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex";
    String w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid";
    String w16="http://schemas.microsoft.com/office/word/2018/wordml";
    String w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash";
    String w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex";
    String wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
    String wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk";
    String wne="http://schemas.microsoft.com/office/word/2006/wordml";
    String wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    String Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14";

    List<WordItem> wordItems = new ArrayList<>();
    SectPr sectPr = new SectPr(PaperType.A4);

    public DocumentContent(){}

    public DocumentContent(Element element){
        Element body = element.element("body");
        List<Element> ps = body.elements("p");
        for(Element p :ps){
            Element r = p.element("r");
            Element drawing = r.element("drawing");
            if(drawing != null){
                WordImage wordImage = new WordImage(drawing);
            }
        }

        System.out.printf("");
    }

    public CoreProperties toCoreProperties(){
        CoreProperties doc = new CoreProperties("w","document");
        CoreProperties body = new CoreProperties("w","body");

        doc.addAttribute("xmlns:wpc",wpc);
        doc.addAttribute("xmlns:cx",cx);
        doc.addAttribute("xmlns:cx1",cx1);
        doc.addAttribute("xmlns:cx2",cx2);
        doc.addAttribute("xmlns:cx3",cx3);
        doc.addAttribute("xmlns:cx4",cx4);
        doc.addAttribute("xmlns:cx5",cx5);
        doc.addAttribute("xmlns:cx6",cx6);
        doc.addAttribute("xmlns:cx7",cx7);
        doc.addAttribute("xmlns:cx8",cx8);

        doc.addAttribute("xmlns:mc",mc);
        doc.addAttribute("xmlns:aink",aink);
        doc.addAttribute("xmlns:am3d",am3d);
        doc.addAttribute("xmlns:o",o);
        doc.addAttribute("xmlns:r",r);
        doc.addAttribute("xmlns:m",m);
        doc.addAttribute("xmlns:wp14",wp14);
        doc.addAttribute("xmlns:wp",wp);
        doc.addAttribute("xmlns:w10",w10);

        doc.addAttribute("xmlns:w",w);
        doc.addAttribute("xmlns:w14",w14);
        doc.addAttribute("xmlns:w15",w15);
        doc.addAttribute("xmlns:w16cex",w16cex);
        doc.addAttribute("xmlns:w16cid",w16cid);
        doc.addAttribute("xmlns:w16",w16);
        doc.addAttribute("xmlns:w16sdtdh",w16sdtdh);
        doc.addAttribute("xmlns:w16se",w16se);
        doc.addAttribute("xmlns:wpg",wpg);

        doc.addAttribute("xmlns:wpi",wpi);
        doc.addAttribute("xmlns:wne",wne);
        doc.addAttribute("xmlns:wps",wps);
        doc.addAttribute("mc",Ignorable);
        if(wordItems.size() > 0){
            wordItems.forEach(e->{
                body.addChild(e.toCoreProperties());
            });
        }
        body.addChild(sectPr.toCoreProperties());
        doc.addChild(body);
        return doc;
    }

}
