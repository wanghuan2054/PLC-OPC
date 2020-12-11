package dde;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;

public class DateUtils {
    private static final String SDF_Format = "yyyyMMddHHmmss";
    public static String getEventTimekey() {
        DateFormat df = new SimpleDateFormat(SDF_Format);
        return df.format(Calendar.getInstance().getTime());
    }
}
