package com.example;

import java.sql.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Test3 {

    private static final String DRIVER = "com.sap.db.jdbc.Driver";
    private static final String URL = "jdbc:sap://172.30.6.10:30013/?databaseName=QAS";

    private static final String USER_NAME = "ETL";
    private static final String PASSWORD = "P@ssw0rd";

    public Test3() {

    }

    public static void main(String[] args) {
        Long tableSize=-1L;
        if(args.length==1){
            System.out.println("全部导出");
        }else if(args.length==2){
            System.out.println("部分导出");
            try {
                if(args[1]!= null ) {
                    tableSize = Long.valueOf(args[1]);
                }
            } catch (NumberFormatException e) {
                e.printStackTrace();
            }
        }else {
            System.out.println("参数有误");
            return;
        }
        String tableName = args[0];

        Test3 demo = new Test3();
        try {
              demo.select(tableName,tableSize);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void select(String tableName,Long tableSize) throws Exception {
        Connection con = this.getConnection(DRIVER, URL, USER_NAME, PASSWORD);
        if(tableSize>0L){
            PreparedStatement pstmt = con.prepareStatement("select * from "+tableName+" limit "+tableSize+" offset 0");
//            PreparedStatement pstmt = con.prepareStatement("SELECT  Name  FROM  sys.Databases");
            ResultSet rs = pstmt.executeQuery();
            try {
                this.processResult(rs,tableName);
            } finally {
                this.closeConnection(con, pstmt);
            }
        }else {
            PreparedStatement pstmt = con.prepareStatement("select * from " + tableName+" limit 10000 offset 0 ");
            ResultSet rs = pstmt.executeQuery();
            try {
                this.processResult(rs,tableName);
            } finally {
                this.closeConnection(con, pstmt);
            }
        }


    }

    private void processResult(ResultSet rs,String tableName) throws Exception {

        if (rs.next()) {
            ResultSetMetaData rsmd = rs.getMetaData();
            int colNum = rsmd.getColumnCount();

            List<String> colNameList = new ArrayList<String>();
            for (int i = 1; i <= colNum; i++) {
//                System.out.println(rsmd.getColumnName(i));
                colNameList.add(rsmd.getColumnName(i));
                if(i==colNum){
                    System.out.println("一共"+i+"列");
                }
            }

            System.out.print("\n");
            System.out.println("———————–");

            Map<String, List<String>> map = new HashMap<String, List<String>>();
            do {
                ArrayList<String> datas = new ArrayList<String>();
                for (int i = 1; i <= colNum; i++) {
                    datas.add(rs.getString(i) == null ? "" : rs
                            .getString(i).trim());
                }
                map.put(datas.get(0).toString(),datas);

            } while (rs.next());
            //调用EXCEL工具类
            System.out.print("\n");
            System.out.println("———————–");
//            System.out.println(map.toString());
            DBToExcel.createExcel(map,colNameList,tableName);
        } else {
            System.out.println("query not result.");
        }

    }

    private Connection getConnection(String driver, String url, String user,
                                     String password) throws Exception {
        Class.forName(driver);
        return DriverManager.getConnection(url, user, password);

    }

    private void closeConnection(Connection con, Statement stmt)
            throws Exception {
        if (stmt != null) {
            stmt.close();
        }
        if (con != null) {
            con.close();
        }
    }

}
