package sunyu.util.test;

import org.junit.jupiter.api.Test;
import org.ttzero.excel.entity.Workbook;
import sunyu.util.AppendHeaderWorksheetWriter;
import sunyu.util.AppendKeyMapSheet;

import java.io.IOException;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class TestEec {
    @Test
    void t003() throws IOException {
        new Workbook()
                .addSheet(new AppendKeyMapSheet<Object>() {
                    int page = 1;

                    @Override
                    protected List<Map<String, Object>> more() {
                        return getRows(page++);
                    }
                }.setSheetWriter(new AppendHeaderWorksheetWriter()))
                .writeTo(Paths.get("d:/tmp"));//最终想要的结果是A/B/C/D列都有，如果超出范围自动分页
    }

    List<Map<String, Object>> getRows(int page) {
        if (page > 2) {
            return null;
        }
        List<Map<String, Object>> rows = new ArrayList<>();//模拟hbase中的数据，这里查询了hbase
        if (page == 1) {
            for (int i = 0; i < 1000000; i++) {
                int finalI = i;
                if (i <= 500000) {
                    rows.add(new HashMap<String, Object>() {{
                        put("A", "a" + page + finalI);
                    }});
                } else {
                    rows.add(new HashMap<String, Object>() {{
                        put("A", "a" + page + finalI);
                        put("B", "b" + page + finalI);
                    }});
                }
            }
        } else if (page == 2) {
            for (int i = 0; i < 1000000; i++) {
                int finalI = i;
                if (i <= 500000) {
                    rows.add(new HashMap<String, Object>() {{
                        put("A", "a" + page + finalI);
                    }});
                } else {
                    rows.add(new HashMap<String, Object>() {{
                        put("A", "a" + page + finalI);
                        put("B", "b" + page + finalI);
                    }});
                }
            }
        }
        return rows;
    }
}
