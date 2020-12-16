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
        Test3 demo = new Test3();
        try {
            demo.select();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void select() throws Exception {
        Connection con = this.getConnection(DRIVER, URL, USER_NAME, PASSWORD);
        PreparedStatement pstmt = con.prepareStatement("select * from SYS.USERS");
        ResultSet rs = pstmt.executeQuery();
        try {
            this.processResult(rs);
        } finally {
            this.closeConnection(con, pstmt);
        }

    }

    private void processResult(ResultSet rs) throws Exception {

        if (rs.next()) {
            ResultSetMetaData rsmd = rs.getMetaData();
            int colNum = rsmd.getColumnCount();

//            String[] strArray = null;
            List<String> colNameList = new ArrayList<String>();
            for (int i = 1; i <= colNum; i++) {
//                if (i == 1) {
//                    System.out.print("第"+i+"行是"+rsmd.getColumnName(i));
//                } else {
//                    System.out.print("\t" +"第"+i+"行是"+ rsmd.getColumnName(i));
//                }
                colNameList.add(rsmd.getColumnName(i));
                if(i==colNum){
                    System.out.println("一共"+i+"行");
                }
            }

            System.out.print("\n");
            System.out.println("———————–");

            Map<String, List<String>> map = new HashMap<String, List<String>>();
            do {
                ArrayList<String> datas = new ArrayList<String>();
                for (int i = 1; i <= colNum; i++) {
//                    if (i == 1) {
//                        System.out.print(rs.getString(i));
//                    } else {
//                        System.out.print("\t"
//                                + (rs.getString(i) == null ? "" : rs
//                                .getString(i).trim()));
//                    }
                    datas.add(rs.getString(i) == null ? "" : rs
                            .getString(i).trim());
                }
                map.put(datas.get(0).toString(),datas);

            } while (rs.next());
            //调用EXCEL工具类
            System.out.print("\n");
            System.out.println("———————–");
            System.out.println(map.toString());
            DBToExcel.createExcel(map,colNameList);
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
