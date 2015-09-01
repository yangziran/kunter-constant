/**
 * 
 */
package cn.kunter.common.constant.entity;

import java.util.ArrayList;
import java.util.List;

/**
 * 表类型
 * @author yangziran
 * @version 1.0 2014年10月20日
 */
public class Table {

    /** 表名称 */
    private String tableName;
    /** 表备注 */
    private String remarks;
    /** 主键集合 */
    private List<Column> primaryKey = new ArrayList<Column>();
    /** 列集合 */
    private List<Column> cols = new ArrayList<Column>();

    /**
     * 取得 tableName
     * @return tableName String
     */
    public String getTableName() {
        return tableName;
    }

    /**
     * 设定 tableName
     * @param tableName String
     */
    public void setTableName(String tableName) {
        this.tableName = tableName;
    }

    /**
     * 取得 remarks
     * @return remarks String
     */
    public String getRemarks() {
        return remarks;
    }

    /**
     * 设定 remarks
     * @param remarks String
     */
    public void setRemarks(String remarks) {
        this.remarks = remarks;
    }

    /**
     * 取得 primaryKey
     * @return primaryKey List<Column>
     */
    public List<Column> getPrimaryKey() {
        return primaryKey;
    }

    /**
     * 设定 primaryKey
     * @param primaryKey List<Column>
     */
    public void setPrimaryKey(List<Column> primaryKey) {
        this.primaryKey = primaryKey;
    }

    public void addPrimaryKey(Column column) {
        primaryKey.add(column);
    }

    /**
     * 取得 cols
     * @return cols List<Column>
     */
    public List<Column> getCols() {
        return cols;
    }

    /**
     * 设定 cols
     * @param cols List<Column>
     */
    public void setCols(List<Column> cols) {
        this.cols = cols;
    }

    public void addCols(Column column) {
        cols.add(column);
    }
}