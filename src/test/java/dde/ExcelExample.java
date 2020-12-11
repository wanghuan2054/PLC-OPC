package dde;

import com.pretty_tools.dde.DDEException;
import com.pretty_tools.dde.DDEMLException;
import com.pretty_tools.dde.client.DDEClientConversation;

/**
 * Excel Data Capture With DDE .
 *
 * @author Wanghuan
 * 循环读取Excel数据内容 ， 每5S
 */
public class ExcelExample {
    public static int ROWS = 100;
    public static int MILLS = 5000;

    public static void main(String[] args) {
        while (true) {
            try {
                // DDE client
                final DDEClientConversation conversation = new DDEClientConversation();
                conversation.setTimeout(3000);
                // Establish conversation with opened and active workbook
                conversation.connect("Excel", "Sheet1");
                // if you have several opened files, you can establish conversation using file path
//          conversation.connect("Excel", "C:\\Users\\10024908\\Desktop\\FlinkCourse\\src\\main\\resources\\FOR_CIM.xlsx");
                try {
                    // Requesting CELL value
                    String colA = null;
                    String colB = null;
                    String colC = null;
                    for (int i = 1; i < ROWS; i++) {
                        colA = "R" + i + "C" + 1;
                        colB = "R" + i + "C" + 2;
                        colC = "R" + i + "C" + 3;
                        if (conversation.request(colA).trim().length() == 0) {
                            continue;
                        } else {
                            System.out.println(conversation.request(colA).trim() + " " +
                                    conversation.request(colB).trim() + " " +
                                    conversation.request(colC).trim());
                        }
                    }
                } finally {
                    conversation.disconnect();
                    Thread.sleep(MILLS);
                }
            } catch (DDEMLException e) {
                System.out.println("DDEMLException: 0x" + Integer.toHexString(e.getErrorCode()) + " " + e.getMessage());
            } catch (DDEException e) {
                System.out.println("DDEClientException: " + e.getMessage());
            } catch (Exception e) {
                System.out.println("Exception: " + e);
            }
        }
    }
}