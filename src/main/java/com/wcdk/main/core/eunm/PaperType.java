package com.wcdk.main.core.eunm;

public enum PaperType {
    A4(11906,16838,1440,1800,1440,1800,851,992,0);
    int w;
    int h;
    int top;
    int right;
    int bottom;
    int left;
    int header;
    int footer;
    int gutter;

    PaperType(int w, int h, int top, int right, int bottom, int left, int header, int footer, int gutter) {
        this.w = w;
        this.h = h;
        this.top = top;
        this.right = right;
        this.bottom = bottom;
        this.left = left;
        this.header = header;
        this.footer = footer;
        this.gutter = gutter;
    }

    public int getW() {
        return w;
    }

    public int getH() {
        return h;
    }

    public int getTop() {
        return top;
    }

    public int getRight() {
        return right;
    }

    public int getBottom() {
        return bottom;
    }

    public int getLeft() {
        return left;
    }

    public int getHeader() {
        return header;
    }

    public int getFooter() {
        return footer;
    }

    public int getGutter() {
        return gutter;
    }
}
