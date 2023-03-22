package com.wcdk.main.word.table;

import com.wcdk.main.word.paragraph.Paragraph;
import com.wcdk.main.word.core.CoreProperties;
import com.wcdk.main.word.core.WordItem;
import lombok.Data;

import java.util.ArrayList;
import java.util.List;

@Data
public class WordTable implements WordItem {
    String elementId;
    String name = "tbl";
    String prefix = "w";
    TblPr tableProportes = new TblPr("5","auto","dxa","autofit");
    int rowNum;
    int cellNum;
    List<Row> rows;
    /** tblGrid gridCol**/
    List<Integer> cellWith = new ArrayList<>();

    public String getElementId() {
        return elementId;
    }

    public void setElementId(String elementId) {
        this.elementId = elementId;
    }

    public WordTable(int rowNum, int cellNum) {
        this.rows = new ArrayList<>();
        for(int i = 0;i < rowNum;i++){
            Row row = new Row();
            row.createCell(cellNum);
            rows.add(row);
        }
       for(int i = 0; i < cellNum;i++){
           cellWith.add(2840);
       }
       this.rowNum = rowNum;
       this.cellNum = cellNum;
    }
    public WordTable() {

    }
    /** cellnum 列数 默认1 小于1 所有列合并行 **/
    public WordTable mergeRow(int cellnum,int start,int end){
        if(cellnum > this.cellNum){
            throw new RuntimeException("列数大于实际列");
        }
        if(start > end){
            throw new RuntimeException("起始入行数大于结束行数");
        }
        if(end > this.rowNum){
            throw new RuntimeException("结束输入行数大于实际行数");
        }
        if(start >= end ){
            throw new RuntimeException("起始输入行数必须小于结束行数");
        }


        if(cellnum >= 1){
            Row row = this.rows.get(start-1);
            Row.Cell cell = row.getCells().get(cellnum-1);
            cell.tcPr.vMerge="restart";

            for(int index = start;index < end;index++){
                Row row_o = this.rows.get(index);
                Row.Cell cell_o = row_o.getCells().get(cellnum-1);
                cell_o.tcPr.vMerge="continue";
                cell.addParagraph(cell_o.getParagraph());
                cell_o.setParagraph(new ArrayList<>());
            }

        }else {
            for(int i =0;i < this.cellNum;i++){
                Row row = this.rows.get(start-1);
                Row.Cell cell = row.getCells().get(i);
                cell.tcPr.vMerge="restart";

                for(int index = start;index < end;index++){
                    Row.Cell cell_o = row.getCells().get(i);
                    cell_o.tcPr.vMerge="continue";
                    cell.addParagraph(cell_o.getParagraph());
//                    cell_o.setParagraph(new ArrayList<>());
                }
            }
        }
        return this;
    }
    /** rownum 行数 默认1 小于1 所有行合并列 **/
    public WordTable mergeCell(int rownum,int start,int end){
        if(rownum > this.rowNum){
            throw new RuntimeException("起始输入行数大于实际行数");
        }
        if(start > this.cellNum){
            throw new RuntimeException("起始输入列数大于实际列数");
        }
        if(start >= end ){
            throw new RuntimeException("起始入列数必须小于结束列数");
        }
        if(end > this.cellNum){
            throw new RuntimeException("结束输入列数大于实际列数");
        }
        if(rownum >= 1){
            int index = rownum-1;
            Row row = this.rows.get(index);
            List<Row.Cell> cells = row.getCells();
            Row.Cell cell = cells.get(start-1);
            cell.tcPr.gridSpan = end -start+1;
            for(int i = start;i< end;i++){
                Row.Cell cell1 = cells.get(i);
                cell.tcPr.w = (Integer.parseInt(cell.tcPr.w)+Integer.parseInt(cell1.tcPr.w))+"";

                cell.addParagraph(cell1.getParagraph());
                cells.remove(i);
                i--;
                end--;
            }
        }else {
            for(int index = 0;index < this.rowNum;index++){
                Row row = this.rows.get(index);
                List<Row.Cell> cells = row.getCells();
                Row.Cell cell = cells.get(0);
                cell.tcPr.gridSpan = end -start;
                cells.clear();
                cells.add(cell);
            }
        }
        return this;
    }
    @Override
    public CoreProperties toCoreProperties() {
        CoreProperties tbl = new CoreProperties("w","tbl");
        CoreProperties tblPr = this.tableProportes.toCoreProperties();
        CoreProperties tblGrid = new CoreProperties("w","tblGrid");
        for (Integer integer:cellWith){
            CoreProperties gridCol = new CoreProperties("w","gridCol");
            gridCol.addAttribute("w:w",integer.intValue()+"");
            tblGrid.addChild(gridCol);
        }
        tbl.addChild(tblPr,tblGrid);
        for (Row row:rows){
            tbl.addChild(row.toCoreproperties());
        }
        return tbl;
    }

    @Data
    public class TblPr {
        String tblStyle;
        String tblW;
        String tblInd;
        String tblLayout;
        String name = "tblPr";

        Borders borders = new Borders();
        Margins tblCellMar = new Margins("0", "180", "0", "180");

        public TblPr(String tblStyle, String tblW, String tblInd, String tblLayout) {
            this.tblStyle = tblStyle;
            this.tblW = tblW;
            this.tblInd = tblInd;
            this.tblLayout = tblLayout;
        }

        public TblPr(String name) {
            this.name = name;
        }
        public TblPr() {

        }

        public CoreProperties toCoreProperties() {
            CoreProperties tblPr = new CoreProperties("w", this.name);
            if(this.tblStyle != null){
                CoreProperties tblStyle = new CoreProperties("w", "tblStyle");
                tblStyle.addAttribute("w:val", this.tblStyle);
                tblPr.addChild(tblStyle);
            }
           if(tblW != null){
               CoreProperties tblW = new CoreProperties("w", "tblW");
               tblW.addAttribute("w:w", "0");
               tblW.addAttribute("w:type", this.tblW);
               tblPr.addChild(tblW);
           }
           if(tblInd != null){
               CoreProperties tblInd = new CoreProperties("w", "tblInd");
               tblInd.addAttribute("w:w", "0");
               tblInd.addAttribute("w:type", this.tblInd);
               tblPr.addChild(tblInd);
           }
            if(tblLayout != null){
                CoreProperties tblLayout = new CoreProperties("w", "tblLayout");
                tblLayout.addAttribute("w:type", this.tblLayout);
                tblPr.addChild(tblLayout);
            }
            CoreProperties borders = this.borders.toCoreproperties();
            CoreProperties tblCellMar = this.tblCellMar.toCoreProperties();

            tblPr.addChild(borders, tblCellMar);

            return tblPr;
        }
    }

    private class Margins {
        @Data
        private class Margin {
            String prefix = "w";
            String fixed;
            String type = "dax";
            String w = "0";

            public Margin(String fixed, String w) {
                this.fixed = fixed;
                this.w = w;
            }

            public CoreProperties toCoreProperties() {
                CoreProperties coreProperties = new CoreProperties("w", this.fixed);
                coreProperties.addAttribute("w:w", this.w);
                coreProperties.addAttribute("w:type", "dxa");
                return coreProperties;
            }
        }

        Margin top = new Margin("top", "0");
        Margin left = new Margin("left", "180");
        Margin bottom = new Margin("bottom", "0");
        Margin right = new Margin("right", "180");

        public Margins(String top, String left, String bottom, String right) {
            this.top.setW(top);
            this.left.setW(left);
            this.bottom.setW(bottom);
            this.right.setW(right);
        }

        public CoreProperties toCoreProperties() {
            CoreProperties coreProperties = new CoreProperties("w", "tblCellMar");
            CoreProperties topm = this.top.toCoreProperties();
            CoreProperties leftm = this.left.toCoreProperties();
            CoreProperties bottomm = this.bottom.toCoreProperties();
            CoreProperties rightm = this.right.toCoreProperties();
            coreProperties.addChild(topm, leftm, bottomm, rightm);
            return coreProperties;
        }
    }

    @Data
    public class Borders {
        @Data
        private class Border {
            String attribute_val = "single";
            String attribute_color = "auto";
            String attribute_sz = "4";
            String attribute_space = "0";

            protected CoreProperties toCoreproperties(String name) {
                CoreProperties border = new CoreProperties("w", name);
                border.addAttribute("w:space", this.attribute_space);
                border.addAttribute("w:color", this.attribute_color);
                border.addAttribute("w:val", this.attribute_val);
                border.addAttribute("w:sz", this.attribute_sz);

                return border;
            }

        }

        Border top = new Border();
        Border left = new Border();
        Border bottom = new Border();
        Border right = new Border();
        Border insideH = new Border();
        Border insideV = new Border();

        protected CoreProperties toCoreproperties() {
            CoreProperties tblBorders = new CoreProperties("w", "tblBorders");
            CoreProperties top = this.top.toCoreproperties("top");
            CoreProperties left = this.left.toCoreproperties("left");
            CoreProperties bottom = this.bottom.toCoreproperties("bottom");
            CoreProperties right = this.right.toCoreproperties("right");
            CoreProperties insideH = this.insideH.toCoreproperties("insideH");
            CoreProperties insideV = this.insideV.toCoreproperties("insideV");
            tblBorders.addChild(top, left, bottom, right, insideH, insideV);
            return tblBorders;
        }

    }

    public class Row {
        private List<Cell> cells = new ArrayList<>();
        private TblPr tblPrEx = new TblPr("tblPrEx");

        @Data
        public class Cell {
            private List<WordItem> paragraph = new ArrayList<>();
            private TcPr tcPr = new TcPr();
            @Data
            public class TcPr {
                String w = "2840";
                String type = "dxa";
                int gridSpan;/** 横向 **/
                String vMerge;/**纵向 restart  continue**/
                public CoreProperties toCoreProperties(){
                    CoreProperties tcPr = new CoreProperties("w","tcPr");
                    CoreProperties tcW = new CoreProperties("w","tcW");
                    tcW.addAttribute("w:w",this.w);
                    tcW.addAttribute("w:type",this.type);
                    tcPr.addChild(tcW);
                    if(gridSpan > 0){
                        CoreProperties gridSpan = new CoreProperties("w","gridSpan");
                        gridSpan.addAttribute("w:val",this.gridSpan+"");
                        tcPr.addChild(gridSpan);
                    }
                    if(vMerge != null && !vMerge.equals("")){
                        CoreProperties vMerge = new CoreProperties("w","vMerge");
                        vMerge.addAttribute("w:vMerge",this.vMerge);
                        tcPr.addChild(vMerge);
                    }
                    return tcPr;
                }
            }
            public Cell addParagraph(WordItem paragraph){
                this.paragraph.add(paragraph);
                return this;
            }
            public Cell addParagraph(List<WordItem> paragraph){
                this.paragraph.addAll(paragraph);
                return this;
            }
            public CoreProperties toCoreProperties(){
                CoreProperties coreProperties = new CoreProperties("w","tc");
                CoreProperties tcPr = this.tcPr.toCoreProperties();
                coreProperties.addChild(tcPr);
                if(this.paragraph.size() > 0){
                   for(int i = 0; i < this.paragraph.size(); i++){
                       CoreProperties item = this.paragraph.get(i).toCoreProperties();
                       coreProperties.addChild(item);
                   }
                }else {
                    coreProperties.addChild(new Paragraph().toCoreProperties());
                }
                return coreProperties;
            }

        }

        public Row createCell(int num){
            for(int i = 0; i < num;i++){
                Cell cell = new Cell();
                cells.add(cell);
            }
            return this;
        }

        public List<Cell> getCells(){
            return this.cells;
        }
        protected CoreProperties toCoreproperties() {
            CoreProperties tr = new CoreProperties("w","tr");
            CoreProperties tblPrEx = this.tblPrEx.toCoreProperties();
            tr.addChild(tblPrEx);
            for(Cell cell:cells){
                tr.addChild(cell.toCoreProperties());
            }
            return tr;
        }
    }

}
