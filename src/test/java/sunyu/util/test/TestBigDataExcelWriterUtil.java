package sunyu.util.test;

import cn.hutool.core.io.FileUtil;
import cn.hutool.log.Log;
import cn.hutool.log.LogFactory;
import org.junit.jupiter.api.Test;
import sunyu.util.BigDataExcelWriterUtil;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.TreeMap;

public class TestBigDataExcelWriterUtil {
    Log log = LogFactory.get();

    @Test
    void t001() {
        BigDataExcelWriterUtil writerUtil = BigDataExcelWriterUtil.builder()
                .destFile("d:/tmp/1.xlsx").sheetName("Sheet")
                .pageSize(1000000).cacheSize(5000)
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
        log.info("done");
    }

    @Test
    void t002() {
        BigDataExcelWriterUtil writerUtil = BigDataExcelWriterUtil.builder()
                .build(FileUtil.file("d:/tmp/1.xlsx"));
        //模拟多行数据
        for (int i = 0; i < 1000100; i++) {
            Map<String, Object> row = new TreeMap<>();
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
