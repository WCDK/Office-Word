package com.wen.main.word.eunm;

public enum RelationshipType {
    fontTable("http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"),
    customXml("http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"),
    image("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"),
    theme("http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"),
    footer("http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"),
    header("http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"),
    settings("http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"),
    styles("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
    private String value;

    RelationshipType(String value) {
        this.value =value;
    }
}
