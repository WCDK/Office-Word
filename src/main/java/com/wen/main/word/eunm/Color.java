package com.wen.main.word.eunm;

public enum Color {
    red("FF0000"),
    yellow("FFC000"),
    green("00B050"),
    blue("0000FF");
    String code;
    Color(String value) {
        this.code =value;
    }

    public String getCode() {
        return code;
    }

    @Override
    public String toString() {
        return this.code;
    }
}
