# Excel工具类

## 描述

* 为了解决动态列大数据导出，不固定列的问题，扩展BigExcelWriter类
* 因为Hbase是列式存储，有可能每一行的列都不同，那么导出excel的时候，表头就是动态的
* 可以按行读取超大excel文件，避免内存溢出
* 目前只适用于一行表头

## 环境

* jdk8 x64 及以上版本

## 依赖

```xml

<dependency>
    <groupId>sunyu.util</groupId>
    <artifactId>util-excel</artifactId>
    <!-- {util.version}_{jdk.version}_{architecture.version} -->
    <version>1.0_jdk8_x64</version>
</dependency>
```

## 例子

```java
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
        //单独追加一行，只有一列的
        writerUtil.append(new LinkedHashMap<String, Object>() {{
            put("列名2", "单独追加的列值");
        }});
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

}


@Test
void t002() {
    String filePath = "d:/tmp/从23年11月开始截止到现在已激活无流量卡(只有武汉和天盛对接的数据)2.xlsx";
    BigDataExcelReaderUtil bigDataExcelReaderUtil = BigDataExcelReaderUtil.builder()
            .setFilePath(filePath)//读取文件路径
            //.setFile(file)//读取文件
            .setRid(-1)//读取所有sheet；-1表示读取全部Sheet, 0表示只读取第一个Sheet
            .build();
    log.info("{}", bigDataExcelReaderUtil.getSheetNames());
    log.info("{}", bigDataExcelReaderUtil.getSheetHeaders());
    bigDataExcelReaderUtil.read(excelRow -> {
        // 处理ExcelRow对象
        log.info("{} {} {} {} {}",
                excelRow.getSheetIndex(),
                excelRow.getSheetName(),
                excelRow.getRowIndex(),
                excelRow.getRowMap(),
                excelRow.getRowCells()
        );
    });
    bigDataExcelReaderUtil.close();
}
```

