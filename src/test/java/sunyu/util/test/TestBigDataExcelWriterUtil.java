package sunyu.util.test;

import cn.hutool.log.Log;
import cn.hutool.log.LogFactory;
import org.junit.jupiter.api.Test;
import sunyu.util.BigDataExcelWriterUtil;

import java.util.LinkedHashMap;
import java.util.Map;

public class TestBigDataExcelWriterUtil {
    Log log = LogFactory.get();

    @Test
    void t001() {
        BigDataExcelWriterUtil writerUtil = BigDataExcelWriterUtil.builder()
                .setDestFile("d:/tmp/1.xlsx").setSheetName("Sheet")
                .setPageSize(1000000).setCacheSize(5000)
                .build();
        //模拟多行数据
        for (int i = 0; i < 1000100; i++) {
            Map<String, Object> row = new LinkedHashMap<>();//使用LinkedHashMap来保持列顺序
            row.put("列名1", i);
            row.put("列名2", i);
            row.put("列名3", i);
            writerUtil.append(row);//追加行
        }
        //更改表头别名
        writerUtil.setHeadersAlias(headers -> {
            //将 列名1 改成 列名one
            headers.put("列名1", "列名one");
            //更改其他列名
        });
        //写出文件
        writerUtil.write();
        writerUtil.close();
    }

}
