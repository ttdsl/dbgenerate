package com.kayak.tools.properties;

import com.kayak.tools.buss.GenerateBussiness;

import java.io.FileNotFoundException;
import java.sql.*;
import java.util.Properties;

public class GenerateProperties {

    /**
     * 需要生成的表名用逗号分隔,默认所有
     * T8_ALGORITHM_LOG,t8_bond_deal_info
     */
    public static String generateTables;

    /**
     * 不需要生成的表名用逗号分隔,默认空
     * T8_ALGORITHM_LOG,t8_bond_deal_info
     */
    public static String excludeTables;
    /**
     * 需要生成或者更新SCHEMA的位置
     */
    public static String excelPath = "D:/test/KKWS_SCHEMA.xls";

    /**
     * 数据库连接配置
     */
    private static String jdbcUrl = "jdbc:oracle:thin:@localhost:1521/orcl";

    private static String driverClass = "oracle.jdbc.driver.OracleDriver";

    private static String userName = "fms_cloud";

    private static String passWord = "fms_cloud";

    public static void initProperties(Properties properties)throws FileNotFoundException{
        checkProperties(properties);

        excelPath = String.valueOf(properties.get("generate.excelPath"));
        jdbcUrl = String.valueOf(properties.get("generate.jdbcUrl"));
        driverClass = String.valueOf(properties.get("generate.driverClass"));
        userName = String.valueOf(properties.get("generate.userName"));
        passWord = String.valueOf(properties.get("generate.passWord"));

        excludeTables = String.valueOf(properties.get("generate.exexcludeTables"));
        generateTables = String.valueOf(properties.get("generate.generateTables"));
        if(!GenerateBussiness.isBlank(excludeTables)){
            String [] exclude = excludeTables.split(",");
            StringBuilder tmp = new StringBuilder();
            for(int i=0;i<exclude.length;i++){
                tmp.append("'");
                tmp.append(exclude[i]).append("'");
                if(i!=exclude.length-1){
                    tmp.append(",");
                }
            }
            excludeTables = tmp.toString();
        }
        if(!GenerateBussiness.isBlank(generateTables)){
            String [] exclude = generateTables.split(",");
            StringBuilder tmp = new StringBuilder();
            for(int i=0;i<exclude.length;i++){
                tmp.append("'");
                tmp.append(exclude[i]).append("'");
                if(i!=exclude.length-1){
                    tmp.append(",");
                }
            }
            generateTables = tmp.toString();
        }
    }
    private static void checkProperties(Properties properties)throws FileNotFoundException{
        if(properties.containsKey("generate.jdbcUrl")&&properties.containsKey("generate.driverClass")&&
                properties.containsKey("generate.userName")&&properties.containsKey("generate.passWord")){
            return;
        }
        throw new FileNotFoundException("参数不对,前检查 application.properties文件");
    }

    public static Connection getConnection()throws SQLException,ClassNotFoundException,IllegalAccessException,InstantiationException{
        Connection conn =null;
        try {
            Driver driver = (Driver) Class.forName(driverClass).newInstance();

            Properties info = new Properties(); //driver的connect方法中需要一个Properties型的参数
            info.put("user", userName);
            info.put("password", passWord);

            //4.使用driver的connect方法获取数据库连接
            conn = driver.connect(jdbcUrl, info);
            return conn;
        }catch(SQLException e){
            System.out.println("获取数据库连接失败");
            throw new SQLException("获取数据库连接失败");
        }
    }

    public static ResultSet queryData(String sql)throws SQLException,ClassNotFoundException,IllegalAccessException,InstantiationException{
        System.out.println("执行SQL"+sql);
        try {
            Connection connection = getConnection();
            Statement statement = connection.createStatement();
            return statement.executeQuery(sql);
        }catch (SQLException e){
            System.out.println("SQL查询错误"+sql);
            throw new SQLException("SQL查询错误"+sql);
        }
    }

}
