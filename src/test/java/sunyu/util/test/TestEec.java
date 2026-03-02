package sunyu.util.test;

import java.io.IOException;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

import org.junit.jupiter.api.Test;
import org.ttzero.excel.entity.SimpleSheet;
import org.ttzero.excel.entity.Workbook;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.Row;
import org.ttzero.excel.reader.Sheet;

import cn.hutool.log.Log;
import cn.hutool.log.LogFactory;
import sunyu.util.ExcelUtil;
import sunyu.util.config.AppendHeaderWorksheetWriter;
import sunyu.util.config.AppendKeyMapSheet;
import sunyu.util.test.pojo.FaHuo;

public class TestEec {
    Log log = LogFactory.get();
    ExcelUtil excelUtil = ExcelUtil.builder().build();

    @Test
    void test_read() {
        List<Map<String, Object>> list = excelUtil.read(Paths.get("d:/tmp/发货明细/20260227/20260226发货明细.xlsx"), 0, 1, 1);
        for (Map<String, Object> m : list) {
            log.info("{}", m);
            log.info("{}", m.get("工况号"));
        }
    }

    @Test
    void test_read2() {
        excelUtil.read(Paths.get("d:/tmp/excel/2026016发货明细.xlsx"), excelReader -> excelReader
                .sheet(0)
                .asFullSheet()
                .copyOnMerged() // <- 转为FullSheet并复制合并单元格
                .header(1, 1)
                .rows()
                .map(row -> row.too(FaHuo.class))
                .forEach(new Consumer<FaHuo>() {
                    @Override
                    public void accept(FaHuo faHuo) {
                        log.info("{} {} {}", faHuo.getIccid(), faHuo.getSimMsisdn(), faHuo.getExprTime());
                    }
                }));
    }

    @Test
    void 写出一个Sheet() {
        // 准备导出数据
        List<Object> rows = new ArrayList<>();
        rows.add(new String[] { "列1", "列2", "列3" });
        rows.add(new int[] { 1, 2, 3, 4 });
        rows.add(new Object[] { 5, new Date(), 7, null, "字母", 9, 10.1243 });
        excelUtil.write("test", Paths.get("d:/tmp"), new SimpleSheet<Object>(rows));
    }

    /**
     * 打印所有worksheet的内容如果有多个sheet页时，因为我们调用了sheets()方法，
     * 此方法会返回一个Stream<Sheet>它会带出所有Sheet页
     */
    @Test
    void 读取一个文件() {
        try (ExcelReader reader = ExcelReader.read(Paths.get("d:/tmp/excel/2026016发货明细.xlsx"))) {
            reader.sheets().flatMap(Sheet::rows).forEach(System.out::println);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void 读取第一个Sheet() {
        try (ExcelReader reader = ExcelReader.read(Paths.get("d:/tmp/excel/2026016发货明细.xlsx"))) {
            // 按行读取第1个Sheet并打印
            reader.sheet(0).rows().forEach(System.out::println);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void 读取第一个Sheet2() {
        try (ExcelReader reader = ExcelReader.read(Paths.get("d:/tmp/excel/2026016发货明细.xlsx"))) {
            // 按行读取第1个Sheet并打印
            reader
                    .sheet(0)// 获取第1个Sheet页
                    .dataRows()// 读取第一个非空行做为表头解析
                    .forEach(System.out::println);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void 读取第一个Sheet为map() {
        try (ExcelReader reader = ExcelReader.read(Paths.get("d:/tmp/excel/2026016发货明细.xlsx"))) {
            reader
                    .sheet(0)
                    .asFullSheet()
                    .copyOnMerged() // <- 转为FullSheet并复制合并单元格
                    .header(1)
                    .rows()
                    .map(Row::toMap)
                    .forEach(System.out::println);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void 大数据量写动态表头() throws IOException {
        new Workbook()
                .addSheet(new AppendKeyMapSheet<Object>() {
                    int page = 1;

                    @Override
                    protected List<Map<String, Object>> more() {
                        return getRows(page++);
                    }
                }.setSheetWriter(new AppendHeaderWorksheetWriter()))
                .writeTo(Paths.get("d:/tmp"));
    }

    List<Map<String, Object>> getRows(int page) {
        if (page > 2) {
            return null;
        }
        List<Map<String, Object>> rows = new ArrayList<>();
        if (page == 1) {
            for (int i = 0; i < 1000000; i++) {
                int finalI = i;
                if (i <= 500000) {
                    rows.add(new HashMap<String, Object>() {
                        {
                            put("A", "a" + page + finalI);
                        }
                    });
                } else {
                    rows.add(new HashMap<String, Object>() {
                        {
                            put("A", "a" + page + finalI);
                            put("B", "b" + page + finalI);
                        }
                    });
                }
            }
        } else if (page == 2) {
            for (int i = 0; i < 1000000; i++) {
                int finalI = i;
                if (i <= 500000) {
                    rows.add(new HashMap<String, Object>() {
                        {
                            put("A", "a" + page + finalI);
                        }
                    });
                } else {
                    rows.add(new HashMap<String, Object>() {
                        {
                            put("A", "a" + page + finalI);
                            put("B", "b" + page + finalI);
                        }
                    });
                }
            }
        }
        return rows;
    }
}