package sunyu.util.test;

import cn.hutool.log.Log;
import cn.hutool.log.LogFactory;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.WorkbookUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;
import sunyu.util.BigDataExcelReaderUtil;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class TestBigDataExcelReaderUtil {
    private static final Log log = LogFactory.get();

    @Test
    void t001() {
        String filePath = "d:/tmp/从23年11月开始截止到现在已激活无流量卡(只有武汉和天盛对接的数据)2.xlsx";
        List<String> sheetNames = new ArrayList<>();
        // 使用Map存储每个sheet的标题
        Map<Integer, List<String>> sheetHeaders = new HashMap<>();
        // 1. 预先获取所有Sheet名称
        try {
            Workbook workbook = WorkbookUtil.createBook(filePath);
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheetNames.add(workbook.getSheetName(i));
            }
            workbook.close(); // 关闭Workbook避免文件占用
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        ExcelUtil.readBySax(filePath, -1, (sheetIndex, rowIndex, rowCells) -> {
            if (rowIndex == 0) {
                // 将标题行转换为String类型并去除空格
                List<String> headers = new ArrayList<>();
                for (Object cell : rowCells) {
                    headers.add(cell.toString().trim());  // 标题去除左右空格
                }
                sheetHeaders.put(sheetIndex, headers);
            } else {
                // 获取当前sheet的标题
                List<String> headers = sheetHeaders.get(sheetIndex);
                if (headers != null) {
                    // 将行数据转换为Map，值保持Object类型
                    Map<String, Object> rowMap = new HashMap<>();
                    for (int i = 0; i < headers.size(); i++) {
                        if (i < rowCells.size()) {
                            Object value = rowCells.get(i);
                            // 如果是String类型，去除左右空格
                            if (value instanceof String) {
                                value = ((String) value).trim();
                            }
                            rowMap.put(headers.get(i), value);
                        }
                    }
                    log.info("{} {} {}", sheetNames.get(sheetIndex), rowIndex, rowMap);
                }
            }
        });
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
}
