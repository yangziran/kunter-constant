/**
 * 
 */
package cn.kunter.common.constant.entity;

/**
 * 列类型
 * @author yangziran
 * @version 1.0 2014年10月20日
 */
public class Column {

    private String serial;
    private String remarks;
    private String columnName;
    private String sqlType;
    private Integer length;
    
    /**
     * 取得 serial
     * @return serial String
     */
    public String getSerial() {
        return serial;
    }
    
    /**
     * 设定 serial
     * @param serial String
     */
    public void setSerial(String serial) {
        this.serial = serial;
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
     * 取得 columnName
     * @return columnName String
     */
    public String getColumnName() {
        return columnName;
    }
    
    /**
     * 设定 columnName
     * @param columnName String
     */
    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }
    
    /**
     * 取得 sqlType
     * @return sqlType String
     */
    public String getSqlType() {
        return sqlType;
    }
    
    /**
     * 设定 sqlType
     * @param sqlType String
     */
    public void setSqlType(String sqlType) {
        this.sqlType = sqlType;
    }
    
    /**
     * 取得 length
     * @return length Integer
     */
    public Integer getLength() {
        return length;
    }
    
    /**
     * 设定 length
     * @param length Integer
     */
    public void setLength(Integer length) {
        this.length = length;
    }

}