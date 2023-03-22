package com.wcdk.main.word.paragraph;

import com.wcdk.main.word.core.CoreProperties;
import com.wcdk.main.word.core.WordItem;
import com.wcdk.main.word.core.eunm.Algin;
import lombok.Data;

@Data
public class Paragraph implements WordItem {
    String hint="eastAsia";
    String eastAsiaTheme = "minorEastAsia";
    String eastAsia = "zh-CN";
    WordItem wordItem;
    String text;
    String fontColor;
    String style;
    String bidi;
    Algin algin;
    String highlight;
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
                CoreProperties bidi = new CoreProperties("w","bidi").addAttribute("w:val","0");
                pPr.addChild(pStyle,bidi);
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
