package com.kayak.tools.bean;

public class ColumnInfo{
    private String is_pk;
    private String column_name;
    private String data_type;
    private String data_length;
    private String nullable;
    private String data_default;
    private String column_comment;

    public String getIs_pk() {
        return is_pk;
    }

    public void setIs_pk(String is_pk) {
        this.is_pk = is_pk;
    }

    public String getColumn_name() {
        return column_name;
    }

    public void setColumn_name(String column_name) {
        this.column_name = column_name;
    }

    public String getData_type() {
        return data_type;
    }

    public void setData_type(String data_type) {
        this.data_type = data_type;
    }

    public String getData_length() {
        return data_length;
    }

    public void setData_length(String data_length) {
        this.data_length = data_length;
    }

    public String getNullable() {
        return nullable;
    }

    public void setNullable(String nullable) {
        this.nullable = nullable;
    }

    public String getData_default() {
        return data_default;
    }

    public void setData_default(String data_default) {
        this.data_default = data_default;
    }

    public String getColumn_comment() {
        return column_comment;
    }

    public void setColumn_comment(String column_commnet) {
        this.column_comment = column_commnet;
    }

    @Override
    public String toString() {
        return "ColumnInfo{" +
                "is_pk='" + is_pk + '\'' +
                ", column_name='" + column_name + '\'' +
                ", data_type='" + data_type + '\'' +
                ", data_length='" + data_length + '\'' +
                ", nullable='" + nullable + '\'' +
                ", data_default='" + data_default + '\'' +
                ", column_commnet='" + column_comment + '\'' +
                '}';
    }
}
