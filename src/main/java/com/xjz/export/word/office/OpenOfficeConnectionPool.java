package com.xjz.export.word.office;

import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;

import java.util.LinkedList;

/**
 * @author: Gibbons
 * @create: 2018/06/08 19:02
 * @description:
 **/
public class OpenOfficeConnectionPool {

    //初始化连接数目
    private int initCount = 3;

    //最大连接数目
    private int initMaxCount = 6;

    //记录当前使用的连接数目
    private int currentCount = 0;

    //连接池   使用linkedList集合存储   存放初始化的连接
    private LinkedList<OpenOfficeConnection> connectionPool = new LinkedList<OpenOfficeConnection>();

    //构造函数  初始化连接
    public OpenOfficeConnectionPool(int initCount) {
        //初始化initCount个连接
        for (int i = 0; i < initCount; i++) {
            connectionPool.add(createConnection());
        }
    }

    /**
     * 创建连接的方法
     *
     * @return OpenOfficeConnection
     */
    public OpenOfficeConnection createConnection() {
        try {
            return new SocketOpenOfficeConnection("211.149.224.37", 8100);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 获取连接的方法
     *
     * @return OpenOfficeConnection
     */
    public OpenOfficeConnection getConnection() {
        //判断，池中若有连接直接使用
        if (connectionPool.size() > 0) {
            //把这个链接移出集合并返回当前连接对象。
            currentCount++;
            return connectionPool.removeFirst();
        }
        //如果池中没有连接而且没有达到最大连接数目；则创建连接
        if (currentCount >= initCount && currentCount < initMaxCount) {
            currentCount++;
            return createConnection(); //创建一个新的连接
        }
        //判断是否达到最大连接数，达到则抛出异常
        throw new RuntimeException("当前连接已经达到最大连接数！");
    }

    /**
     * 释放连接的方法
     *
     * @param conn
     */
    public void releaseConnection(OpenOfficeConnection conn) {
        //判断池中的数目如果小于初始化连接就放回连接池中。
        //判断连接池中的剩余数目是否<连接池初始化数目   如果为真 则放回连接池
        if (currentCount <= initCount) {
            connectionPool.addLast(conn);  //放回连接池
            currentCount--; //当前连接-1
        } else {
            conn.disconnect(); //关闭连接
            currentCount--;
        }
    }

    public static void main(String[] args) {
        OpenOfficeConnectionPool pool = new OpenOfficeConnectionPool(3);

        OpenOfficeConnection conn1 = pool.getConnection();
        OpenOfficeConnection conn2 = pool.getConnection();
        OpenOfficeConnection conn3 = pool.getConnection();
        OpenOfficeConnection conn4 = pool.getConnection();

        System.out.println("池中初始化的连接:" + pool.initCount);
        System.out.println("最大连接数目:" + pool.initMaxCount);

        System.out.println("池中剩余连接:" + pool.connectionPool.size());
        System.out.println("当前连接数目:" + pool.currentCount);
    }
}
