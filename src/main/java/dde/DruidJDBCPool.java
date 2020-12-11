package dde;


import com.alibaba.druid.pool.DruidDataSource;

import java.sql.Connection;

/**
 * 数据库连接生成类，返回一个数据库连接对象
 * 构造函数完成数据库驱动的加载和数据的连接
 * 提供数据库连接的取得和数据库的关闭方法
 * @author wanghuan
 *
 */

public class DruidJDBCPool {


    private static DruidJDBCPool instance;
    private DruidJDBCPool(){

    };
    public static DruidJDBCPool  getInstance(){
        if(instance==null){
            synchronized (DruidJDBCPool.class){
                if(instance==null){
                    instance=new DruidJDBCPool();
                }
            }
        }return instance;
    }
    private static DruidDataSource dataSource=null;

    /**
     * 构造函数完成数据库的连接和连接对象的生成
     * @throws Exception
     */

    public void GetDbConnect() throws Exception  {
        try{
            if(dataSource==null){
                dataSource=new DruidDataSource();
                //设置连接参数
                dataSource.setUrl("jdbc:oracle:thin:@10.120.8.20:1521:MDWDB1");
                dataSource.setDriverClassName("oracle.jdbc.OracleDriver");
                dataSource.setUsername("edbadm");
                dataSource.setPassword("edbadm");
                //配置初始化大小、最小、最大
                dataSource.setInitialSize(1);
                dataSource.setMinIdle(1);
                dataSource.setMaxActive(30);
                //连接泄漏监测
                dataSource.setRemoveAbandoned(true);
                dataSource.setRemoveAbandonedTimeout(300);
                //配置获取连接等待超时的时间
                dataSource.setMaxWait(20000);
                //配置间隔多久才进行一次检测，检测需要关闭的空闲连接，单位是毫秒
                dataSource.setTimeBetweenEvictionRunsMillis(20000);
                //防止过期
                dataSource.setValidationQuery("SELECT 'x' FROM DUAL");
                dataSource.setTestWhileIdle(true);
                dataSource.setTestOnBorrow(true);
            }
        }catch(Exception e){
            throw e;
        }
    }

    /**
     * 取得已经构造生成的数据库连接
     * @return 返回数据库连接对象
     * @throws Exception
     */
    public Connection getConnect() throws Exception{
        Connection con=null;
        try {
            GetDbConnect();
            con=dataSource.getConnection();
        } catch (Exception e) {
            throw e;
        }
        return con;
    }
}
