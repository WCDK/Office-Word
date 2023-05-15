package com.wcdk.main.excel;

import com.wcdk.main.core.CoreProperties;
import com.wcdk.main.core.OfficeItem;
import lombok.Data;

import java.util.ArrayList;
import java.util.List;

public class Sheet implements OfficeItem {
    String name;
    int sheetId = 1;
    String rId = "rId"+sheetId;
    /**dimension start:endrows**/
    String start = "A1";
    String end = "A1";
    String rows = "1";

    String xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    String xmlns_r="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    String xmlns_xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    String xmlns_x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    String xmlns_mc="http://schemas.openxmlformats.org/markup-compatibility/2006";
    String xmlns_etc="http://www.wps.cn/officeDocument/2017/etCustomData";

    List<SheetData> sheetDatas = new ArrayList<>();

    boolean tabSelected = false;
    String workbookViewId = "0";
    /** sqref**/
    String activeCell = "A1";
    String defaultColWidth = "9";
    String defaultRowHeight="13.5";
    String min = "1";
    String max = "16384";
    String width = "9";
    String style = "1";

    @Override
    public CoreProperties toCoreProperties() {
        CoreProperties worksheet = new CoreProperties(null,"worksheet");
        worksheet.addAttribute("xmlns",xmlns)
                .addAttribute("xmlns:r",xmlns_r)
                .addAttribute("xmlns:xdr",xmlns_xdr)
                .addAttribute(" xmlns:x14",xmlns_x14)
                .addAttribute("xmlns:mc",xmlns_mc)
                .addAttribute("xmlns:etc",xmlns_etc);

        CoreProperties dimension = new CoreProperties(null,"dimension");
        dimension.addAttribute("ref",start+":"+end+rows);

        CoreProperties sheetViews = new CoreProperties(null,"sheetViews");
        CoreProperties sheetView = new CoreProperties(null,"sheetView");
        if(tabSelected){
            sheetView.addAttribute("tabSelected","1");
        }
        sheetView.addAttribute("workbookViewId",workbookViewId);
        CoreProperties selection = new CoreProperties(null,"selection");
        selection.addAttribute("activeCell",activeCell)
                .addAttribute("sqref",activeCell);
        CoreProperties sheetFormatPr = new CoreProperties(null,"sheetFormatPr");
        sheetFormatPr.addAttribute("defaultColWidth",defaultColWidth)
                .addAttribute("defaultRowHeight",defaultRowHeight)
                .addAttribute("outlineLevelRow","2")
                .addAttribute("outlineLevelCol","2");
        sheetView.addChild(selection);
        sheetViews.addChild(sheetView);

        CoreProperties cols = new CoreProperties(null,"cols");
        CoreProperties col = new CoreProperties(null,"col");
        col.addAttribute("minx",min)
                .addAttribute("max",max)
                .addAttribute("width",width)
                .addAttribute("style",style);
        cols.addChild(col);

        CoreProperties pageMargins = new CoreProperties(null,"pageMargins");
        pageMargins.addAttribute("left","0.75")
                .addAttribute("right","0.75")
                .addAttribute("top","1")
                .addAttribute("bottom","0.5")
                .addAttribute("footer","0.5");
        worksheet.addChild(dimension,sheetViews,sheetFormatPr,cols);
        sheetDatas.forEach(e->worksheet.addChild(e.toCoreProperties()));
        worksheet.addChild(pageMargins);
    return worksheet;
    }
    public CoreProperties getSlot() {
        CoreProperties sheet = new CoreProperties(null,"sheet");
        sheet.addAttribute("name",name)
                .addAttribute("sheetId",sheetId+"")
                .addAttribute("r:id","rId"+sheetId);

        return sheet;
    }
    @Data
    class SheetData{
        String r;
        String spans;
        List<C> rows = new ArrayList<>();
        @Data
       class C{
           String r;
           String s;
           String t;
           String  v;
            public CoreProperties toCoreProperties() {
                CoreProperties c = new CoreProperties(null,"c");
                c.addAttribute("r",r)
                        .addAttribute("s",s)
                        .addAttribute("t","t");

                CoreProperties v = new CoreProperties(null,"v",this.v);
                c.addChild(v);
                return c;
            }
       }
       public CoreProperties toCoreProperties() {
            CoreProperties row = new CoreProperties(null,"row");
            row.addAttribute("r",r)
                    .addAttribute("spans",spans);
            rows.forEach(e->row.addChild(e.toCoreProperties()));
            return row;
        }
    }
}
