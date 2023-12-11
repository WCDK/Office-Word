package com.wcdk.main.core;

import lombok.Data;

@Data
public class Style {
    String type;
    String styleId;

    String name;
    String basedOn;
    String next;
    String qFormat;
    String uiPriority;
    public Style(){

    }
}
