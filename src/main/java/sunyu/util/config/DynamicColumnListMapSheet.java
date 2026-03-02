package sunyu.util.config;

import java.util.Arrays;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import org.ttzero.excel.entity.Column;
import org.ttzero.excel.entity.ListMapSheet;
import org.ttzero.excel.reader.Cell;

/**
 * 动态列的ListMapSheet
 * 
 * 一般用于列头数据不是固定的
 *
 * @author SunYu
 */
public class DynamicColumnListMapSheet<T> extends ListMapSheet<T> {
    // 保存已存在中列
    Set<String> existsKeys = new HashSet<>();

    @Override
    public Column[] getHeaderColumns() {
        Column[] columns = super.getHeaderColumns();
        for (Column col : columns) {
            existsKeys.add(col.name); // <- 初始化表头装载到existsKeys
        }
        return columns;
    }

    @Override
    protected void resetBlockData() {
        if (!eof && left() < rowBlock.capacity())
            append();
        int end = getEndIndex(), len;
        for (; start < end; rows++, start++) {
            org.ttzero.excel.entity.Row row = rowBlock.next();
            row.index = rows;
            row.height = getRowHeight();
            Map<String, ?> rowDate = data.get(start);
            boolean isNull = rowDate == null;

            if (!isNull) {
                // 检查是否有不存在的key
                for (String k : rowDate.keySet()) {
                    if (!existsKeys.contains(k)) {
                        existsKeys.add(k); // <- 将Key添加到existsKeys
                        Column col = new Column(k, k), pre = columns[columns.length - 1].getTail(); // <- 需要判断NPE
                        col.colIndex = pre.colIndex + 1;
                        col.colNum = pre.getColNum() + 1;
                        col.styles = getWorkbook().getStyles();
                        // 扩容并追加到末尾
                        columns = Arrays.copyOf(columns, columns.length + 1);
                        columns[columns.length - 1] = col;
                    }
                }
            }
            len = columns.length;

            Cell[] cells = row.realloc(len);
            for (int i = 0; i < len; i++) {
                Column hc = columns[i];
                Object e = !isNull ? rowDate.get(hc.key) : null;
                // Clear cells
                Cell cell = cells[i];
                cell.clear();

                cellValueAndStyle.reset(row, cell, e, hc);
            }
        }
    }
}