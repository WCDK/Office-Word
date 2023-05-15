package com.wcdk.main.core.eunm;

public enum Algin {
    left("left"),
    right("right"),
    center("center");
    String value;
    Algin(String value) {
        this.value =value;
    }

    public String getCode() {
        return value;
    }

    @Override
    public String toString() {
        return this.value;
    }
}
