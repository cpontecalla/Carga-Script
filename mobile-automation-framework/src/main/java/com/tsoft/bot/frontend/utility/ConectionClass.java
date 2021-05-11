package com.tsoft.bot.frontend.utility;

import cucumber.deps.com.thoughtworks.xstream.io.copy.HierarchicalStreamCopier;

import java.sql.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Objects;

public class ConectionClass {

        private static Statement sentencia;
        private static ResultSet resultado;

    private static Connection dbConnection() throws SQLException, ClassNotFoundException {
        Connection connection;
        String host = "sl-us-south-1-portal.47.dblayer.com";
        String port = "22832";
        String dataBase = "compose";
        String username = "admin";
        String password = "PHOHIIACWMQRLJBY";
        try {
            Class.forName("org.postgresql.Driver");
            connection = DriverManager.getConnection("jdbc:postgresql://" + host + ":" + port + "/" + dataBase + "","" + username + "","" + password + "");
            if (connection != null){
                System.out.println("Conexion Establecida");
            }else {
                System.out.println("Conexion Fallida");
            }
        }
        catch (Exception we)
        {
            System.out.println("Error: " + we);
            throw new SQLException("");
        }
        return connection;
    }

//    public static void main(String[] args) throws SQLException, ClassNotFoundException {
//        //dbConnection();
//        String token = executeQuerySelect("select token from ibmx_a07e6d02edaf552.tdp_token_vendedor where codatis = '111534';");
//        System.out.println("El token real is: " + token);
//        //closeConnection();
//    }

    public static String executeQuerySelect(String query) throws SQLException, ClassNotFoundException {
        Connection connection = null;
        Statement stmt = null;
        String token = "";
        try {
            connection = dbConnection();
            stmt = connection.createStatement();
            ResultSet rs = stmt.executeQuery(query);
            int columns = rs.getMetaData().getColumnCount();
            while (rs.next()) {
                for (int i = 1; i <= columns; ++i) {
//                    System.out.println("["+md.getColumnName(i)+":   "+rs.getString(md.getColumnName(i))+"]");
                    //row.put(md.getColumnName(i), rs.getString(md.getColumnName(i)));
                    token = ((rs.getObject(i) == null) ? "": rs.getObject(i).toString());
                    System.out.println("Token is: " + token);
                }
                //mydata.add(row);
            }
        } catch (SQLException | ClassNotFoundException e) {
            System.out.println("[Conn-SQL] - ExecuteQuery:  "+e.getMessage());
            throw e;
        }finally {
            if(!Objects.isNull(connection)) closeConnection(connection);
            if(!Objects.isNull(stmt)) stmt.close();
        }
        return token;
    }
    private static void closeConnection(Connection connection) throws SQLException {
        try {
            connection.close();
        } catch (SQLException e) {
            System.out.println("[Conn-SQL] closeConnectionDB - Error al cerrar la Conexión con la Base de Datos" + e.getMessage());
            throw e;
        }
    }
}
