/**
 * 
 */
package cn.kunter.common.constant.make;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import cn.kunter.common.constant.db.ConnectionFactory;
import cn.kunter.common.constant.entity.Column;
import cn.kunter.common.constant.entity.Table;

/**
 * 获取到数据
 * @author 阳自然
 * @version 1.0 2015年9月1日
 */
public class GetData {

    /**
     * 获取到数据
     * @param table
     * @return
     * @throws SQLException
     * @author 阳自然
     */
    public static List<Map<String, Object>> getData(Table table) throws SQLException {

        // 获取数据库连接
        Connection connection = ConnectionFactory.getConnection();
        Statement statement = connection.createStatement();
        // 执行SQL
        ResultSet resultSet = statement
                .executeQuery("select * from " + table.getTableName() + " ORDER BY const_id, serial_number");

        // 遍历结果集
        List<Map<String, Object>> dataList = new ArrayList<>();
        while (resultSet.next()) {
            Map<String, Object> data = new HashMap<>();
            for (Column column : table.getCols()) {
                data.put(column.getColumnName(), resultSet.getObject(column.getColumnName()));
            }
            dataList.add(data);
        }

        resultSet.close();
        statement.close();
        connection.close();

        return dataList;
    }
}
