package com.wcdk.main.word.paragraph;

import com.wcdk.main.core.CoreProperties;
import com.wcdk.main.core.OfficeItem;
import com.wcdk.main.core.eunm.Algin;
import lombok.Data;

@Data
public class Paragraph implements OfficeItem {
    String hint="eastAsia";
    String eastAsiaTheme = "minorEastAsia";
    String eastAsia = "zh-CN";
    OfficeItem wordItem;
    String text;
    String fontColor;
    String style;
    String bidi;
    boolean sort = false;
    Algin algin;
    String highlight;
    String sortIlvl;
    String numId = "1";
    @Override
    public CoreProperties toCoreProperties() {
        if (wordItem != null || text != null){
            CoreProperties p = new CoreProperties("w","p");

            CoreProperties pPr = new CoreProperties("w","pPr");
            CoreProperties rPr = new CoreProperties("w","rPr");
            CoreProperties rFonts = new CoreProperties("w","rFonts");
            rFonts.addAttribute("w:hint",this.hint);
            rFonts.addAttribute("w:eastAsiaTheme",this.eastAsiaTheme);
            CoreProperties lang = new CoreProperties("w","lang").addAttribute("w:eastAsia",this.eastAsia);
            rPr.addChild(rFonts,lang);
            pPr.addChild(rPr);
            CoreProperties r = new CoreProperties("w","r");
            r.addChild(rPr);

            if(this.algin != null){
                CoreProperties jc = new CoreProperties("w","jc").addAttribute("w:val",this.algin.getCode());
                pPr.addChild(jc);
            }
            if(this.fontColor != null){
                CoreProperties color = new CoreProperties("w","color").addAttribute("w:val",this.fontColor);
                pPr.addChild(color);
                rPr.addChild(color);
            }
            if(this.highlight != null){
                CoreProperties highlight = new CoreProperties("w","highlight").addAttribute("w:val",this.highlight);
                pPr.addChild(highlight);
                rPr.addChild(highlight);
            }
            if(this.style != null){
                CoreProperties pStyle = new CoreProperties("w","pStyle").addAttribute("w:val",this.style);
                pPr.addChild(pStyle);
            }
            if(this.bidi != null){
                CoreProperties bidi = new CoreProperties("w","bidi").addAttribute("w:val",this.bidi);
                CoreProperties ind = new CoreProperties("w","ind").addAttribute("w:firstLineChars","0");
                ind.addAttribute("w:left","1151").addAttribute("w:hanging","1151");
                pPr.addChild(bidi,ind);
//                if(sort){
//                    CoreProperties numPr = new CoreProperties("w","numPr");
//                    CoreProperties ilvl = new CoreProperties("w","ilvl").addAttribute("w:val",this.sortIlvl);
//                    CoreProperties numId = new CoreProperties("w","numId").addAttribute("w:val",this.numId);
//                    numPr.addChild(ilvl,numId);
//                    pPr.addChild(numPr);
//                }
            }

            if(wordItem != null){
                r.addChild(wordItem.toCoreProperties());
            }else {
                CoreProperties t = new CoreProperties("w","t");
                t.setValue(text);
                r.addChild(t);
            }
            p.addChild(pPr,r);
            return p;
        }

        return null;
    }
}
