package sunyu.util.pojo;

import java.util.List;
import java.util.Map;

public class ExcelRow {
    /**
     * sheet索引
     */
    private int sheetIndex;
    /**
     * 行索引
     */
    private long rowIndex;
    /**
     * 行数据map格式，key是表头，value是单元格数据
     */
    private Map<String, Object> rowMap;
    /**
     * 行数据list格式，list的索引对应表头索引
     */
    private List<Object> rowCells;

    public ExcelRow(int sheetIndex, long rowIndex, Map<String, Object> rowMap, List<Object> rowCells) {
        this.sheetIndex = sheetIndex;
        this.rowIndex = rowIndex;
        this.rowMap = rowMap;
        this.rowCells = rowCells;
    }

    public int getSheetIndex() {
        return sheetIndex;
    }

    public void setSheetIndex(int sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public long getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(long rowIndex) {
        this.rowIndex = rowIndex;
    }

    public Map<String, Object> getRowMap() {
        return rowMap;
    }

    public void setRowMap(Map<String, Object> rowMap) {
        this.rowMap = rowMap;
    }

    public List<Object> getRowCells() {
        return rowCells;
    }

    public void setRowCells(List<Object> rowCells) {
        this.rowCells = rowCells;
    }
}
