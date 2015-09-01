/**
 * 
 */
package cn.kunter.common.constant.make;

import java.util.List;
import java.util.Map;

import cn.kunter.common.constant.config.PropertyHolder;
import cn.kunter.common.constant.entity.Table;
import cn.kunter.common.constant.util.FileUtil;
import cn.kunter.common.constant.util.OutputUtilities;

/**
 * 根据常量数据生成常量类
 * @author 阳自然
 * @version 1.0 2015年8月28日
 */
public class MakeConstant {

    private final static String PACKAGES = PropertyHolder.getConfigProperty("package");

    public static void main(String[] args) throws Exception {

        List<Table> tables = GetTableConfig.getTableConfig();

        MakeConstant.makerEntity(tables.get(0));
    }

    /**
     * 常量类生成
     * @param table
     * @throws Exception
     * @author yangziran
     */
    public static void makerEntity(Table table) throws Exception {

        StringBuilder builder = new StringBuilder();
        // 包结构
        builder.append("package ").append(PACKAGES).append(";");
        OutputUtilities.newLine(builder);

        OutputUtilities.newLine(builder);
        builder.append("/**");
        OutputUtilities.newLine(builder);
        builder.append(" * 类名称：系统常量定义类");
        OutputUtilities.newLine(builder);
        builder.append(" * @author 工具生成");
        OutputUtilities.newLine(builder);
        builder.append(" * @version 1.0 2015年1月1日");
        OutputUtilities.newLine(builder);
        builder.append(" */");
        OutputUtilities.newLine(builder);

        // 类开始
        builder.append("public abstract interface SystemConstant {");
        OutputUtilities.newLine(builder);

        // 字段定义
        List<Map<String, Object>> dataList = GetData.getData(table);
        String idKey = null;
        for (Map<String, Object> map : dataList) {
            String constId = (String) map.get("const_id");
            String constKey = (String) map.get("const_key");
            String constName = (String) map.get("const_name");
            String constValue = (String) map.get("const_value");
            if (!constId.equals(idKey)) {
                idKey = constId;
                OutputUtilities.newLine(builder);
                builder.append("    /** ").append(constName).append(" */");
                OutputUtilities.newLine(builder);
                builder.append("    public static final String ID_").append(constId).append(" = \"").append(constId)
                        .append("\";");
            }
            OutputUtilities.newLine(builder);
            builder.append("    /** ").append(constName).append(" ").append(constKey).append("：").append(constValue)
                    .append(" */");
            OutputUtilities.newLine(builder);
            builder.append("    public static final String KEY_").append(constId).append("_").append(constKey)
                    .append(" = \"").append(constKey).append("\";");
        }
        OutputUtilities.newLine(builder);
        // 类结束
        builder.append("}");

        // 输出文件
        FileUtil.writeFile(
                PropertyHolder.getConfigProperty("target") + PACKAGES.replaceAll("\\.", "/") + "/SystemConstant.java",
                builder.toString());
    }
}