package com.wcdk.main.excel.xl;

import com.wcdk.main.core.CoreProperties;
import com.wcdk.main.core.OfficeItem;

public class SharedStrings implements OfficeItem {
    String xmsn = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    String count;
    String uniqueCount;
    @Override
    public CoreProperties toCoreProperties() {
        return null;
    }
}
