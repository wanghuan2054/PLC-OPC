package dde;

import com.pretty_tools.dde.DDEException;
import com.pretty_tools.dde.DDEMLException;
import com.pretty_tools.dde.client.DDEClientConversation;
import org.apache.log4j.Logger;

import java.sql.Connection;
import java.sql.PreparedStatement;

/**
 * Excel Data Capture With DDE .
 *
 * @author Wanghuan
 * 循环读取Excel数据内容 ， 每1 min
 */

public class DDEImport {
    public static Logger logger = Logger.getLogger(DDEImport.class);
    /**
     * ROWS 默认采集Excel的行数范围
     * MILLS 数据采集的频率
     */
    public static final int ROWS = 1000;
    public static final int MILLS = 1000*60;
    /**
     * CHECK_EXCEL_ALIVE   Excel 关闭再次检测是否打开的检测时间  30分钟
     * 1000 * 60 * 30
     */
    public static final int CHECK_EXCEL_ALIVE = 1000 * 5;
    // public static final String EXCLE_SHEET_TOPIC = "Sheet1";
    public static final String EXCLE_PATH = "C:\\Users\\10024908\\Desktop\\JDDE\\FOR_CIM.xlsx";
    public static final String ROW_IDENTIIER = "R";
    public static final String COLUMN_IDENTIIER = "C";

    private static Connection connection = null;
    private static PreparedStatement insertPS = null;

    public static void main(String[] args) throws InterruptedException {
        while (true) {
            try {
                // DDE client establish
                final DDEClientConversation conversation = new DDEClientConversation();
                conversation.setTimeout(3000);
                /**
                 *  1. Establish conversation with opened and active workbook
                 *     conversation.connect("Excel", EXCLE_SHEET_TOPIC);
                 *  2. if you have several opened files, you can establish conversation using file path
                 *     conversation.connect("Excel", EXCLE_PATH);
                 */

                conversation.connect("Excel", EXCLE_PATH);
                logger.info("conversation is connecting");

                connection = DruidJDBCPool.getInstance().getConnect();
                connection.setAutoCommit(false);
                try {
                    // Requesting CELL value
                    String colA = null;
                    String colB = null;
                    String colC = null;
                    // 每一次采集，时间点取值一致
                    String event_timekey = DateUtils.getEventTimekey();
                    for (int i = 1; i < ROWS; i++) {
                        colA = ROW_IDENTIIER + i + COLUMN_IDENTIIER + 1;
                        colB = ROW_IDENTIIER + i + COLUMN_IDENTIIER + 2;
                        colC = ROW_IDENTIIER + i + COLUMN_IDENTIIER + 3;
                        if (conversation.request(colA).trim().length() == 0) {
                            continue;
                        } else {
/*                            System.out.println(conversation.request(colA).trim() + " " +
                                    conversation.request(colB).trim() + " " +
                                    conversation.request(colC).trim());*/
                            String insertSql = "insert into EDS_ENERGY_EHS(EQP_ID   ,  \n" +
                                    "ITEM_NAME ,\n" +
                                    "VALUE_NUM,\n" +
                                    "EVENT_TIMEKEY\n" + ") VALUES (?,?,?,?)";

                            insertPS = connection.prepareStatement(insertSql);
                            insertPS.setString(1, conversation.request(colA).trim());
                            insertPS.setString(2, conversation.request(colB).trim());
                            insertPS.setFloat(3, Float.valueOf(conversation.request(colC).trim()));
                            insertPS.setString(4, event_timekey);
                            insertPS.executeUpdate();
                        }
                    }
                    // 每一次采集多行数据，批量提交一次
                    connection.commit();
                    logger.info("power data info collect success ,save success ! ");
                } finally {
                    conversation.disconnect();
                    Thread.sleep(MILLS);
                }
            } catch (DDEMLException e) {
                logger.error("DDEMLException: 0x" + Integer.toHexString(e.getErrorCode()) + " " + e.getMessage());
                // 如果客户端Excel异常关闭，程序每隔CHECK_EXCEL_ALIVE时间探测一次Excel是否打开
                if (16394 == e.getErrorCode()) {
                    logger.error("Excel 已关闭，程序自动" + CHECK_EXCEL_ALIVE + "分钟后进行再次检测Excel是否打开");
                    // 线程阻塞，CHECK_EXCEL_ALIVE 分钟
                    Thread.sleep(CHECK_EXCEL_ALIVE);
                }
            } catch (DDEException e) {
                logger.error("DDEClientException: " + e.getMessage());
            } catch (Exception e) {
                logger.error("Exception: " + e.getMessage());
            }
        }
    }
}