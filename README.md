# dbGenerate

#### 介绍
将Oracle数据库的表结构、表说明、索引等数据查出生成Excel，构建超链接方便使用

#### 软件架构
Java代码编写，单机小程序、使用依赖poi3.8.jar、poi-scratchpad-3.8.jar、class-12.jar、ojdbc14-10.2.jar

#### 使用说明

1.  修改application.properties,配置数据库用户名密码、驱动和JDBC连接以及Excel生成路径(包含Excel名)
2.  代码运行 GenerateDBSchema --> run
3.  jar包运行 java -jar generate.jar
