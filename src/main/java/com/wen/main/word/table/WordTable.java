package com.wen.main.word.table;

import com.wen.main.word.core.CoreProperties;
import com.wen.main.word.core.WordItem;
import lombok.Data;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Data
public class WordTable implements WordItem {
    String name = "tbl";
    String prefix = "w";
    TblPr tableProportes = new TblPr("5","auto","dxa","autofit");
    List<Row> rows;
    /** tblGrid gridCol**/
    List<Integer> cellWith = new ArrayList<>();

    public WordTable(int rowNum,int cellNum) {
        this.rows = new ArrayList<>();
        for(int i = 0;i < rowNum;i++){
            Row row = new Row();
            row.createCell(cellNum);
            rows.add(row);
        }
       for(int i = 0; i < cellNum;i++){
           cellWith.add(2840);
       }
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
    private class TblPr {
        String tblStyle;
        String tblW;
        String tblInd;
        String tblLayout;

        Borders borders = new Borders();
        Margins tblCellMar = new Margins("0", "180", "0", "180");

        public TblPr(String tblStyle, String tblW, String tblInd, String tblLayout) {
            this.tblStyle = tblStyle;
            this.tblW = tblW;
            this.tblInd = tblInd;
            this.tblLayout = tblLayout;
        }

        public TblPr() {
        }

        public CoreProperties toCoreProperties() {
            CoreProperties tblPr = new CoreProperties("w", "tblPr");
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
        private TblPr tblPrEx = new TblPr();

        @Data
        public class Cell {
            private WordItem Paragraph;
            private TcPr tcPr = new TcPr();
            @Data
            public class TcPr {
                String w = "2840";
                String type = "dxa";
            }

            public CoreProperties toCoreProperties(){
                CoreProperties coreProperties = new CoreProperties("w","tc");
                CoreProperties tcPr = new CoreProperties("w","tcPr");
                CoreProperties tcW = new CoreProperties("w","tcW");
                tcW.addAttribute("w:w",this.tcPr.w);
                tcW.addAttribute("w:type",this.tcPr.type);
                tcPr.addChild(tcW);
                coreProperties.addChild(tcPr);
                if(this.Paragraph != null){
                    CoreProperties item = this.Paragraph.toCoreProperties();
                    coreProperties.addChild(item);
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
