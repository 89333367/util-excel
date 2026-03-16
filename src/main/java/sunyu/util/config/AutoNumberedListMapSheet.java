package sunyu.util.config;

import org.ttzero.excel.entity.ListMapSheet;

public class AutoNumberedListMapSheet<T> extends ListMapSheet<T> {
    private String originSheetName;

    @Override
    protected String getCopySheetName() {
        // 第一页（非copy页）
        if (copyCount == 1) {
            originSheetName = name;
            // 默认Sheet特殊处理
            if ("Sheet1".equals(name)) {
                originSheetName = "Sheet1";
            } else {
                name = originSheetName + copyCount;
            }
            copyCount++;
        }
        return originSheetName + copyCount;
    }
}
