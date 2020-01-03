package com.kayak.tools.bean;

import java.util.ArrayList;

public class TableInfo{
    private String table_name;
    private String table_comment;
    private ArrayList<ColumnInfo> columnInfos;
    private ArrayList<IndexInfo> indexInfos;

    public String getTable_name() {
        return table_name;
    }

    public void setTable_name(String table_name) {
        this.table_name = table_name;
    }

    public String getTable_comment() {
        return table_comment;
    }

    public void setTable_comment(String table_comment) {
        this.table_comment = table_comment;
    }

    public ArrayList<ColumnInfo> getColumnInfos() {
        return columnInfos;
    }

    public void setColumnInfos(ArrayList<ColumnInfo> columnInfos) {
        this.columnInfos = columnInfos;
    }

    public ArrayList<IndexInfo> getIndexInfos() {
        return indexInfos;
    }

    public void setIndexInfos(ArrayList<IndexInfo> indexInfos) {
        this.indexInfos = indexInfos;
    }

    @Override
    public String toString() {
        return "TableInfo{" +
                "table_name='" + table_name + '\'' +
                ", column_commnet='" + table_comment + '\'' +
                ", columnInfos=" + columnInfos +
                ", indexInfos=" + indexInfos +
                '}';
    }
}
