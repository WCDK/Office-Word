package com.wcdk.main.excel.xl;

import com.wcdk.main.core.CoreProperties;
import com.wcdk.main.core.OfficeItem;
import com.wcdk.main.excel.Sheet;

import java.util.ArrayList;
import java.util.List;

public class Workbook implements OfficeItem {
    String xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    String xmlns_r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    String appName = "xl";
    String lastEdited="3";
    String lowestEdited = "5";
    String rupBuild = "9302";
    /**bookViews workbookView**/
    int windowWidth = 24225;
    int windowHeight = 12420;
    String calcId = "144525";
    List<Sheet> sheets = new ArrayList<>();
    @Override
    public CoreProperties toCoreProperties() {
        CoreProperties workbook = new CoreProperties(null,"workbook");
        workbook.addAttribute("xmlns",xmlns)
                .addAttribute("xmlns:r",xmlns_r);

        CoreProperties fileVersion = new CoreProperties(null,"fileVersion");
        fileVersion.addAttribute("appName",appName)
                .addAttribute("lastEdited",lastEdited)
                .addAttribute("lowestEdited",lowestEdited)
                .addAttribute("rupBuild",rupBuild);

        CoreProperties bookViews = new CoreProperties(null,"bookViews");
        CoreProperties workbookView = new CoreProperties(null,"workbookView");
        workbookView.addAttribute("windowWidth",windowWidth+"")
                .addAttribute("windowHeight",windowHeight+"");
        bookViews.addChild(workbookView);

        CoreProperties sheete = new CoreProperties(null,"sheets");
        sheets.forEach(e->sheete.addChild(e.getSlot()));

        CoreProperties calcPr = new CoreProperties(null,"calcPr");
        calcPr.addAttribute("calcId",calcId);

        workbook.addChild(fileVersion,bookViews,sheete,calcPr);

        return workbook;
    }
}
