# Excel工具类

## 描述

* 为了解决动态列大数据导出，不固定列的问题，扩展BigExcelWriter类
* 因为Hbase是列式存储，有可能每一行的列都不同，那么导出excel的时候，表头就是动态的
* 注意，这个类非线程安全，一定不要在多线程中使用
* 目前只适用于一行表头

## 环境

* jdk8 x64 及以上版本

## 依赖

```xml

<dependency>
    <groupId>sunyu.util</groupId>
    <artifactId>util-excel</artifactId>
    <version>v1.0</version>
</dependency>
```

## 例子

```java

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
    //回收资源
    writerUtil.close();
}

@Test
void t002() {
    BigDataExcelWriterUtil writerUtil = BigDataExcelWriterUtil.builder()
            .build(FileUtil.file("d:/tmp/1.xlsx"));
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
```

