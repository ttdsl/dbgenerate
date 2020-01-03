package com.kayak.tools.buss;

import com.kayak.tools.bean.ColumnInfo;
import com.kayak.tools.bean.IndexInfo;
import com.kayak.tools.bean.TableInfo;
import com.kayak.tools.properties.GenerateProperties;
import com.kayak.tools.utils.ExcelUtils;
import com.sun.org.apache.xalan.internal.xsltc.compiler.SourceLoader;
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;

import javax.annotation.Resource;
import java.io.*;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Properties;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ForkJoinPool;

/**
 * 根据数据库生成Excel实现
 * @author jintao
 */
public class GenerateBussiness {

    private static HashMap<String,HashMap<String,String>> allTableColumnCommets = new HashMap<>();//存放所有表字段描述
    private static HashMap<String,ArrayList<IndexInfo>> allTableIndexs = new HashMap<String,ArrayList<IndexInfo>>();//存放所有表索引数据
    private static HashMap<String,ArrayList<ColumnInfo>> allTableColumnInfos = new HashMap<String,ArrayList<ColumnInfo>>();//存放所有表字段属性
    public static void run(){

        long start = System.currentTimeMillis();

        try {
            initProperties();

            ArrayList<TableInfo> allTables = new ArrayList<TableInfo>();//存放所有表
            allTables.addAll(queryAllTables());

            allTableColumnCommets = queryAllTableColumnComment();//查询所有字段描述
            allTableIndexs = queryAllTableIndexs();//查询所有索引
            allTableColumnInfos = queryAllTableColumnInfos();//查询所有表字段属性

            allTables.parallelStream().forEach((o) -> {
                o.setIndexInfos(allTableIndexs.get(o.getTable_name()));
                o.setColumnInfos(allTableColumnInfos.get(o.getTable_name()));
            });

            //根据表数据操作Excel
            doCreateOrUpdate(allTables);

            long end = System.currentTimeMillis();
            System.out.println("表说明文件生成完成,耗时"+(end-start)+"毫秒");
        }catch(IOException ioe){
            System.out.println("IO错误");
            ioe.printStackTrace();
        }catch (SQLException sqle){
            System.out.println("SQL错误");
            sqle.printStackTrace();
        }catch (Exception e){
            System.out.println("系统错误");
            e.printStackTrace();
        }
    }
    /**
     * 初始化参数
     */
    private static void initProperties() throws IOException{
        InputStream in = new BufferedInputStream(GenerateBussiness.class.getClassLoader().getResourceAsStream("application.properties"));
        Properties p = new Properties();
        p.load(in);
        GenerateProperties.initProperties(p);
        in.close();
    }
    /**
     * 根据表数据生成或更新表说明文件
     */
    private static void doCreateOrUpdate(ArrayList<TableInfo> allTables)throws IOException,InterruptedException,ExecutionException{
        HSSFWorkbook hssfWorkbook = doCheckPath();//验证配置的EXCEL位置
        FileOutputStream fos = null;
        doCreateOrUpdateIndexSheet(hssfWorkbook,allTables);//判断是否需要更新index页签

        doCreateOrUpdateBodySheet(hssfWorkbook,allTables);//生成或修改其余页签

        doLinkSheet(hssfWorkbook);//将index页签与其他页签加上超链接

        fos = new FileOutputStream(GenerateProperties.excelPath);//生成文件
        hssfWorkbook.write(fos);
        fos.flush();
        fos.close();

    }
    /**
     * 将index上的表名与其余页签加上超链接
     */
    private static void doLinkSheet(HSSFWorkbook hssfWorkbook){
        HSSFSheet hssfSheet = hssfWorkbook.getSheet("index");
        Iterator<Row> rowIterator = hssfSheet.rowIterator();
        while(rowIterator.hasNext()){
            Row row = rowIterator.next();
            Cell cell = row.getCell(0);
            String sheetName = cell.getStringCellValue();
            Hyperlink hyperlink = new HSSFHyperlink(Hyperlink.LINK_DOCUMENT);
            // "#"表示本文档    "sheetName"表示sheet页名称  "A1"表示第一列第一行
            hyperlink.setAddress("#"+sheetName+"!A1");
            cell.setHyperlink(hyperlink);
        }
    }

    /**
     * 创建index Sheet
     */
    private static void doCreateOrUpdateBodySheet(HSSFWorkbook hssfWorkbook,ArrayList<TableInfo> allTables)throws InterruptedException, ExecutionException {
            if (isBlank(GenerateProperties.generateTables)) {
                while(hssfWorkbook.getNumberOfSheets()>2){
                    hssfWorkbook.removeSheetAt(2);
                }
            }
            if(hssfWorkbook.getNumberOfSheets()==2){
                for (int i = 0; i < allTables.size(); i++) {
                    hssfWorkbook.createSheet(allTables.get(i).getTable_name());
                }
            }
            ForkJoinPool forkJoinPool = new ForkJoinPool(5);
            forkJoinPool.submit(() -> {
                allTables.parallelStream().forEach((tableInfo) -> {
                    if (hssfWorkbook.getSheet(tableInfo.getTable_name()) != null) {
                        doUpdateExcelByName(hssfWorkbook,hssfWorkbook.getSheet(tableInfo.getTable_name()), tableInfo);
                    }
                });
            }).get();
    }

    /**
     * 根据表数据更新Excel
     */
    private static void doUpdateExcelByName(HSSFWorkbook hssfWorkbook,HSSFSheet hssfSheet,TableInfo tableInfo) {
        hssfSheet.setColumnWidth(0, 15 * 2 * 256);
        hssfSheet.setColumnWidth(1, 15 * 2 * 256);
        hssfSheet.setColumnWidth(2, 10 * 2 * 256);
        hssfSheet.setColumnWidth(3, 10 * 2 * 256);
        hssfSheet.setColumnWidth(4, 50 * 2 * 256);
        hssfSheet.createFreezePane(0, 1);

        //页签第一行标签处理
        HSSFRow row = hssfSheet.createRow(0);
        row.setHeight((short)350);
        ExcelUtils.cteateCell(hssfWorkbook, row, 0, tableInfo.getTable_name(), ExcelUtils.getCellStyle("head_name"));
        ExcelUtils.cteateCell(hssfWorkbook, row, 1, tableInfo.getTable_comment(), ExcelUtils.getCellStyle("head_name"));

        //页签中表头处理
        HSSFRow row1 = hssfSheet.createRow(1);
        row1.setHeight((short)350);
        ExcelUtils.cteateCell(hssfWorkbook, row1, 0, "Column Is PK", ExcelUtils.getCellStyle("th_left"));
        ExcelUtils.cteateCell(hssfWorkbook, row1, 1, "Column Name", ExcelUtils.getCellStyle("th_left"));
        ExcelUtils.cteateCell(hssfWorkbook, row1, 2, "Column Datatype", ExcelUtils.getCellStyle("th_left"));
        ExcelUtils.cteateCell(hssfWorkbook, row1, 3, "Column Null Option", ExcelUtils.getCellStyle("th_left"));
        ExcelUtils.cteateCell(hssfWorkbook, row1, 4, "Column Comment", ExcelUtils.getCellStyle("th_right"));

        ArrayList<ColumnInfo> columnInfos = tableInfo.getColumnInfos();
        for (int i = 0; i < columnInfos.size(); i++) {
            HSSFRow rowRandom = hssfSheet.createRow(i + 2);
            rowRandom.setHeight((short) 350);
            ColumnInfo columnInfo = columnInfos.get(i);
            CellStyle leftStyle = ExcelUtils.getCellStyle("td_center_left");
            CellStyle rightStyle = ExcelUtils.getCellStyle("td_center_right");
            if (i == columnInfos.size() - 1) {
                leftStyle = ExcelUtils.getCellStyle("td_bottom_left");
                rightStyle = ExcelUtils.getCellStyle("td_bottom_right");
            }
            ExcelUtils.cteateCell(hssfWorkbook, rowRandom, 0, columnInfo.getIs_pk(), leftStyle);
            ExcelUtils.cteateCell(hssfWorkbook, rowRandom, 1, columnInfo.getColumn_name(), leftStyle);
            ExcelUtils.cteateCell(hssfWorkbook, rowRandom, 2, columnInfo.getData_type() + "(" + columnInfo.getData_length() + ")", leftStyle);
            ExcelUtils.cteateCell(hssfWorkbook, rowRandom, 3, ((columnInfo.getNullable().equals("Y") ? "NULL" : "NOT NULL") + (columnInfo.getData_default()==null?"":" DEFAULT"+columnInfo.getData_default())), leftStyle);
            ExcelUtils.cteateCell(hssfWorkbook, rowRandom, 4, columnInfo.getColumn_comment(), rightStyle);
        }
        //插入索引数据
        ArrayList<IndexInfo> indexInfos = tableInfo.getIndexInfos();
        if(indexInfos!=null&&indexInfos.size()>0){
            int start_len = 4+columnInfos.size();
            HSSFRow indexRow = hssfSheet.createRow(start_len);
            start_len++;
            indexRow.setHeight((short) 350);
            CellStyle index = ExcelUtils.getCellStyle("index");
            CellStyle def = ExcelUtils.getCellStyle("def");
            ExcelUtils.cteateCell(hssfWorkbook, indexRow, 0, "Index Name", index);
            ExcelUtils.cteateCell(hssfWorkbook, indexRow, 1, "Index Type", index);
            ExcelUtils.cteateCell(hssfWorkbook, indexRow, 2, "Index Column", index);

            for(int i=0;i<indexInfos.size();i++){
                HSSFRow rowRandom = hssfSheet.createRow(start_len+i);
                rowRandom.setHeight((short) 350);
                IndexInfo indexInfo = indexInfos.get(i);
                ExcelUtils.cteateCell(hssfWorkbook, rowRandom, 0, indexInfo.getIndex_name(), def);
                ExcelUtils.cteateCell(hssfWorkbook, rowRandom, 1, indexInfo.getIndex_type(), def);
                ExcelUtils.cteateCell(hssfWorkbook, rowRandom, 2, indexInfo.getIndex_colmuns(), def);
            }
        }
    }

    /**
     * 创建index Sheet
     */
    private static void doCreateOrUpdateIndexSheet(HSSFWorkbook hssfWorkbook,ArrayList<TableInfo> allTables){
        if(isBlank(GenerateProperties.generateTables)){
            while(hssfWorkbook.getNumberOfSheets()>=2){
                hssfWorkbook.removeSheetAt(1);
            }
            HSSFSheet hssfSheet = hssfWorkbook.createSheet("index");
            hssfSheet.setColumnWidth(0, 20 * 2 * 256);
            hssfSheet.setColumnWidth(1, 20 * 2 * 256);
            for(int i=0;i<allTables.size();i++){
                HSSFRow row = hssfSheet.createRow(i);
                row.setHeight((short)300);
                hssfSheet.createFreezePane(0, 1);
                CellStyle CellStyle = ExcelUtils.getCellStyle("index");
                CellStyle CellStyle1 = ExcelUtils.getCellStyle("def");
                ExcelUtils.cteateCell(hssfWorkbook, row, 0, allTables.get(i).getTable_name(), CellStyle);
                ExcelUtils.cteateCell(hssfWorkbook, row, 1, allTables.get(i).getTable_comment(), CellStyle1);
            }
        }
    }

    /**
     * 验证配置的EXCEL位置
     */
    private static HSSFWorkbook doCheckPath()throws IOException{
        HSSFWorkbook hssfWorkbook = null;
        File file = new File(GenerateProperties.excelPath);
        try {
            if(!file.exists()){
                System.out.println("配置的EXCEL路径不存在,生成文件");
                File parent = new File(file.getParent());
                if(!parent.exists()){
                    parent.mkdirs();
                }
                file.createNewFile();
                hssfWorkbook = new HSSFWorkbook();
                HSSFSheet hssfSheet = hssfWorkbook.createSheet("log");
                hssfSheet.setColumnWidth(0, 100 * 2 * 256);
                HSSFRow row = hssfSheet.createRow(0);
                row.setHeight((short)500);
                hssfSheet.createFreezePane(0, 1);
                CellStyle CellStyle = ExcelUtils.createCellStyle(hssfWorkbook,"LOG");
                ExcelUtils.cteateCell(hssfWorkbook, row, 0, "修改日志记录", CellStyle);
            }else{
                hssfWorkbook = new HSSFWorkbook(new FileInputStream(GenerateProperties.excelPath));
            }
            ExcelUtils.doCreateCellStyle(hssfWorkbook);//生成Cell样式,全局调用,每次重新生成多了Excel样式多了会丢失
        }catch (IOException ioe){
            System.out.println("生成文件失败"+ioe.getMessage());
            throw ioe;
        }
        return hssfWorkbook;
    }
    /**
     * 查询所有表字段属性
     * @return
     */
    private static HashMap<String,ArrayList<ColumnInfo>>  queryAllTableColumnInfos()throws Exception {
        String sql = "select t.table_name,t.column_name,t.data_type,t.data_length,t.nullable,t.data_default from user_tab_columns t where 1=1 ";
        sql = getFullSQL(sql);

        ResultSet resultSet = GenerateProperties.queryData(sql);
        while (resultSet.next()) {
            String table_name = resultSet.getString("table_name");
            String column_name = resultSet.getString("column_name");
            String data_type = resultSet.getString("data_type");
            String data_length = resultSet.getString("data_length");
            String nullable = resultSet.getString("nullable");
            String data_default = resultSet.getString("data_default");
            ColumnInfo columnInfo = new ColumnInfo();
            columnInfo.setColumn_name(column_name);
            columnInfo.setNullable(nullable);
            columnInfo.setData_default(data_default);
            columnInfo.setData_length(data_length);
            columnInfo.setData_type(data_type);

            //判断当前字段是否存在索引
            String  is_pk = "NO";
            ArrayList<IndexInfo> indexs = allTableIndexs.get(table_name);
            if(indexs!=null&&indexs.size()>0){
                a: for(int i=0;i<indexs.size();i++){
                    String index_column = indexs.get(i).getIndex_colmuns();
                    String index_type = indexs.get(i).getIndex_type();
                    String index_name = indexs.get(i).getIndex_name();
                    String [] indexcolumns = index_column.split(",");
                    b : for(String indexColumn:indexcolumns){
                        if(column_name.equals(index_column)){
                            is_pk = "PK";
                            break a;
                        }
                    }
                }
            }
            columnInfo.setIs_pk(is_pk);

            //处理当前字段描述
            String columnComment = allTableColumnCommets.get(table_name).get(column_name);
            columnInfo.setColumn_comment(columnComment);
            if(allTableColumnInfos.get(table_name)!=null){
                allTableColumnInfos.get(table_name).add(columnInfo);
            }else{
                ArrayList<ColumnInfo> columnInfos = new ArrayList<>();
                columnInfos.add(columnInfo);
                allTableColumnInfos.put(table_name,columnInfos);
            }
        }

        return allTableColumnInfos;
    }
    /**
     * 查询所有需要生成表的索引
     * @return
     */
    private static HashMap<String,ArrayList<IndexInfo>>  queryAllTableIndexs()throws Exception{
        String sql = "select max(indexs.table_name)table_name,max(indexs.INDEX_NAME)index_name,max(indexs.UNIQUENESS)uniqueness,wm_concat(indexcol.COLUMN_NAME) column_name from user_indexes indexs left join user_ind_columns indexcol on indexs.INDEX_NAME = indexcol.INDEX_NAME where indexcol.COLUMN_NAME is not null";
        if(!isBlank(GenerateProperties.generateTables)){
            sql += " and indexs.TABLE_NAME in （" + GenerateProperties.generateTables + ")";
        }else if(!isBlank(GenerateProperties.excludeTables)){
            sql += " and indexs.TABLE_NAME not in (" + GenerateProperties.excludeTables + ")";
        }
        sql +=" group by indexs.INDEX_NAME ";
        HashMap<String,ArrayList<IndexInfo>> indexInfos = new HashMap<>();
        try {
            ResultSet resultSet = GenerateProperties.queryData(sql);
            while (resultSet.next()) {
                String table_name = resultSet.getString("table_name");
                String index_name = resultSet.getString("index_name");
                String index_type = resultSet.getString("uniqueness");
                String index_column = resultSet.getString("column_name");
                IndexInfo indexInfo = new IndexInfo();
                indexInfo.setIndex_colmuns(index_column);
                indexInfo.setIndex_name(index_name);
                indexInfo.setIndex_type(index_type);
                if(indexInfos.get(table_name)!=null){
                    indexInfos.get(table_name).add(indexInfo);
                }else{
                    ArrayList<IndexInfo> infos = new ArrayList<>();
                    infos.add(indexInfo);
                    indexInfos.put(table_name,infos);
                }
            }
        }catch(Exception e){
            System.out.println("查询表索引数据失败"+sql);
           throw new Exception("查询表索引数据失败"+sql);
        }
        return indexInfos;
    }

    /**
     * 查询需要生成表字段描述
     * @return
     */
    private static HashMap<String,HashMap<String,String>> queryAllTableColumnComment(){
        String sql = "select t.* from user_col_comments t where 1=1 ";
        sql = getFullSQL(sql);
        HashMap<String,HashMap<String,String>> hs = new HashMap<String,HashMap<String,String>>();
        try {
            ResultSet resultSet = GenerateProperties.queryData(sql);
            while (resultSet.next()) {
                String table_name = resultSet.getString("table_name");
                String column_name = resultSet.getString("column_name");
                String column_comment = resultSet.getString("comments");
                if(hs.get(table_name)!=null){
                    hs.get(table_name).put(column_name,column_comment);
                }else{
                    HashMap<String,String> cols = new HashMap<String,String>();
                    cols.put(column_name,column_comment);
                    hs.put(table_name,cols);
                }
            }
        }catch(Exception e){
            System.out.println("查询表字段描述失败"+sql);
            e.printStackTrace();
        }
        return hs;
    }

    /**
     * 查询需要生成表的表说明
     * @return
     */
    private static ArrayList<TableInfo> queryAllTables(){
        String sql = "select * from user_tab_comments t where 1=1";
        sql = getFullSQL(sql);
        sql += " order by t.table_name";
        ArrayList<TableInfo> tableInfos = new ArrayList<>();
        try {
            ResultSet resultSet = GenerateProperties.queryData(sql);
            while (resultSet.next()) {
                TableInfo tableInfo = new TableInfo();
                tableInfo.setTable_name(resultSet.getString("table_name"));
                tableInfo.setTable_comment(resultSet.getString("comments"));
                ArrayList<ColumnInfo> ar = new ArrayList<>();
                tableInfo.setColumnInfos(ar);
                tableInfos.add(tableInfo);
            }
        }catch(Exception e){
            System.out.println("查询表说明失败"+sql);
            e.printStackTrace();
        }
        return tableInfos;
    }

    /**
     * 查询所有需要生成的表
     * @return
     */
    private static ArrayList queryAllTable(){
        String sql = "select table_name from user_tables t where 1=1";
        sql = getFullSQL(sql);
        ArrayList arrayList = new ArrayList();
        try {
            ResultSet resultSet = GenerateProperties.queryData(sql);
            while (resultSet.next()) {
                arrayList.add(resultSet.getString("table_name"));
            }
        }catch(Exception e){
            System.out.println("查询需要生成的表失败"+sql);
            e.printStackTrace();
        }
        return arrayList;
    }

    private static String getFullSQL(String sql){
        if(!isBlank(GenerateProperties.generateTables)){
            sql += " and t.TABLE_NAME in (" + GenerateProperties.generateTables + ")";
        }else if(!isBlank(GenerateProperties.excludeTables)){
            sql += "and t.TABLE_NAME not in (" + GenerateProperties.excludeTables + ")";
        }
        return sql;
    }

    public static boolean isBlank(String str){
        return str==null||"".equals(str.trim())||"null".equals(str);
    }
}
