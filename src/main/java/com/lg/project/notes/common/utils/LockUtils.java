package com.lg.project.notes.common.utils;

import lombok.Builder;
import lombok.Getter;
import lombok.Setter;

import java.sql.*;
import java.util.Objects;
import java.util.UUID;
import java.util.concurrent.TimeUnit;

/**
 * @ClassName LockUtils
 * @Description 使用mysql实现分布式锁
 * @Author liuguang
 * @Date 2022/5/25 11:26
 * @Version 1.0
 */

public class LockUtils {
    /**
     * 我们创建一个分布式锁表，如下
     * 分布式锁工具类：
     * DROP DATABASE IF EXISTS javacode2018;
     * CREATE DATABASE javacode2018;
     * USE javacode2018;
     * DROP TABLE IF EXISTS t_lock;
     * create table t_lock(
     * lock_key varchar(32) PRIMARY KEY NOT NULL COMMENT '锁唯一标志',
     * request_id varchar(64) NOT NULL DEFAULT '' COMMENT '用来标识请求对象的',
     * lock_count INT NOT NULL DEFAULT 0 COMMENT '当前上锁次数',
     * timeout BIGINT NOT NULL DEFAULT 0 COMMENT '锁超时时间',
     * version INT NOT NULL DEFAULT 0 COMMENT '版本号，每次更新+1'
     * )
     * COMMENT '锁信息表';
     */
    //将requestId保存在该变量中
    private static ThreadLocal<String> requestIdTL = new ThreadLocal<>();
    private static final String url = "jdbc:mysql://localhost:3306/javacode2018? useSSL=false"; //数据库地址
    private static final String username = "root"; //数据库用户名
    private static final String password = "123456"; //数据库密码
    private static final String driver = "com.mysql.jdbc.Driver"; //mysql 驱动

    /**
     * 获取当前线程requestId
     * @return
     */
    public static String getRequestId(){
        String requestId = requestIdTL.get();
        if(requestId == null || "".equals(requestId)){
            requestId = UUID.randomUUID().toString();
            requestIdTL.set(requestId);
        }
        return requestId;
    }

    /**
     * 连接数据库
     * @return
     */
    public static Connection getConn(){
        Connection conn = null;
        try{
            Class.forName(driver); //加载数据库驱动
            conn = DriverManager.getConnection(url,username,password); //连接数据库
        } catch (ClassNotFoundException | SQLException e) {
            e.printStackTrace();
        }
        return conn;
    }

    /**
     * 关闭数据库连接
     * @param conn
     */
    public static void closeConn(Connection conn){
        if(conn != null){
            try {
                conn.close();
            } catch (SQLException e) { //关闭数据库连接
                e.printStackTrace();
            }
        }
    }

    @Getter
    @Setter
    @Builder
    public static class LockModel{
        private String lockKey;
        private String requestId;
        private Integer lockCount;
        private Long timeOut;
        private Integer version;
    }

    public static LockModel get(String lockKey) throws Exception{
        return exec(connection -> {
            String sql = "select * from t_lock t WHERE t.lock_key=?";
            PreparedStatement ps = connection.prepareStatement(sql);
            int colIndex = 1;
            ps.setString(colIndex++,lockKey);
            ResultSet rs = ps.executeQuery();
            if(rs.next()){
                return LockModel.builder().
                        lockKey(lockKey).
                        requestId(rs.getString("request_id")).
                        lockCount(rs.getInt("lock_count")).
                        timeOut(rs.getLong("timeout")).
                        version(rs.getInt("version")).build();
            }
            return null;
        });
    }

    public static int insert(LockModel lockModel) throws Exception {
        return exec(conn -> {
            String sql = "insert into t_lock (lock_key, request_id, lock_count, timeout, version) VALUES (?,?,?,?,?)";
            PreparedStatement ps = conn.prepareStatement(sql);
            int colIndex = 1;
            ps.setString(colIndex++, lockModel.getLockKey());
            ps.setString(colIndex++, lockModel.getRequestId());
            ps.setInt(colIndex++, lockModel.getLockCount());
            ps.setLong(colIndex++, lockModel.getTimeOut());
            ps.setInt(colIndex++, lockModel.getVersion());
            return ps.executeUpdate();
        });
    }

    public static int update(LockModel lockModel) throws Exception {
        return exec(conn -> { String sql = "UPDATE t_lock SET request_id = ?,lock_count = ?,timeout = ?,version = version + 1 WHERE lock_key = ? AND version = ?";
            PreparedStatement ps = conn.prepareStatement(sql);
            int colIndex = 1;
            ps.setString(colIndex++, lockModel.getRequestId());
            ps.setInt(colIndex++, lockModel.getLockCount());
            ps.setLong(colIndex++, lockModel.getTimeOut());
            ps.setString(colIndex++, lockModel.getLockKey());
            ps.setInt(colIndex++, lockModel.getVersion());
            return ps.executeUpdate();
        });
    }

    public static <T> T exec(SqlExec<T> sqlExec) throws Exception{
        Connection conn = getConn();
        try {
            return sqlExec.exec(conn);
        }finally {
            closeConn(conn);
        }
    }

    @FunctionalInterface
    public interface SqlExec<T> {
        T exec(Connection connection) throws Exception;
    }

    /**
     * 获取锁
     * @param lockKey 锁key
     * @param lockTimeOut(毫秒) 持有锁的有效时间，防止死锁
     * @param getTimeOut 获取锁的超时时间，这个时间内获取不到将重试
     * @return
     */
    public static boolean lock(String lockKey,long lockTimeOut,int getTimeOut) throws Exception{
        //默认没有获取到锁
        boolean lockResult = false;
        String requestId = getRequestId();
        long startTime =System.currentTimeMillis();
        while (true){
            LockModel lockModel = LockUtils.get(lockKey);
            if(Objects.isNull(lockModel)){
                //插入一条记录,重新尝试获取锁
                LockUtils.insert(LockModel.builder().lockKey(lockKey).
                        requestId("").lockCount(0).timeOut(0L).version(0).build());
            }else{
                String reqId = lockModel.getRequestId();
                //如果reqId为空字符串,表示锁未被占用
                if("".equals(reqId)){
                    lockModel.setRequestId(requestId);
                    lockModel.setLockCount(1);
                    lockModel.setTimeOut(System.currentTimeMillis()+lockTimeOut);
                    if(LockUtils.update(lockModel) == 1){
                        lockResult = true;
                        break;
                    }

                }else if (requestId.equals(reqId)){
                    //如果request_id和表中request_id一样表示锁被当前线程持有者，此时需要 加重入锁
                    lockModel.setTimeOut(System.currentTimeMillis() + lockTimeOut);
                    lockModel.setLockCount(lockModel.getLockCount() + 1);
                    if(LockUtils.update(lockModel) == 1){
                        lockResult = true;
                        break;
                    }
                }else{
                    //锁不是自己的，并且已经超时了，则重置锁，继续重试
                    if(lockModel.getTimeOut() < System.currentTimeMillis()){
                        //重置锁
                        LockUtils.resetLock(lockModel);
                    }else{
                        //如果未超时，休眠100毫秒，继续重试
                        if(startTime + getTimeOut > System.currentTimeMillis()){
                            TimeUnit.MILLISECONDS.sleep(100);
                        }else{
                            break;
                        }
                    }
                }

            }
        }
        return lockResult;
    }

    /**
     * 重置锁
     * @param lockModel
     * @return
     * @throws Exception
     */
    public static int resetLock(LockModel lockModel) throws Exception{
        lockModel.setRequestId("");
        lockModel.setLockCount(0);
        lockModel.setTimeOut(0L);
        return LockUtils.update(lockModel);
    }

    /**
     * 释放锁
     * @param lockKey
     * @throws Exception
     */
    public static void unlock(String lockKey) throws Exception{
        //获取当前线程requestId
        String requestId = getRequestId();
        LockModel lockModel = LockUtils.get(lockKey);
        //当前线程requestId和库中request_id一致 && lock_count>0，表示可以释放锁
        if(Objects.nonNull(lockModel)
                && requestId.equals(lockModel.getRequestId())
        && lockModel.getLockCount() > 0){
            if(lockModel.getLockCount() > 1){
                //重置锁
                resetLock(lockModel);
            }else{
                lockModel.setLockCount(lockModel.getLockCount() - 1);
                LockUtils.update(lockModel);
            }
        }
    }
}
