package com.wcdk.main.word.core.eunm;

public enum RelationshipType {

    base("http://schemas.openxmlformats.org/package/2006/relationships"),
    fontTable("http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"),
    endnotes("http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"),
    footnotes("http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"),
    customXml("http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"),
    image("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"),
    theme("http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"),
    footer("http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"),
    header("http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"),
    settings("http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"),
    webSettings("http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"),
    styles("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
    private String value;

    RelationshipType(String value) {
        this.value =value;
    }

    public String getValue(){
        return this.value;
    }
    public static RelationshipType getTyppe(String value){
        RelationshipType[] values = RelationshipType.values();
        for(RelationshipType relationshipType : values){
            if(relationshipType.value.equals(value)){
                return relationshipType;
            }
        }
        return null;
    }
}
