package com.sobey.util;

import org.springframework.util.StringUtils;

/**
 * Created by lijunhong on 17/3/10.
 * 主要用于判断前台显示的内容
 */
public class ExcelViewJudge {

    private static final String STORAGESTATE_ONLINE = "0";        //文件状态在线
    private static final String STORAGESTATE_INTO_ONLINE = "1";    //进线状态
    private static final String STORAGESTATE_OFFONLINE = "2";      //离线状态

    private static final String SECRET_LEVEL_PUBLIC = "public";          //公开
    private static final String SECRET_LEVEL_PRIVATE = "private";        //私有的


    /**
     * 根据前台返回的code发送显示的中文
     * @param matching
     * @return
     */
    public static String getPrivilegeTemplateCodeString(String matching){
        if(!StringUtils.isEmpty(matching)){
            if(SECRET_LEVEL_PUBLIC.equalsIgnoreCase(matching)){
                return "公开";
            }
            if(SECRET_LEVEL_PRIVATE.equalsIgnoreCase(matching)){
                return "私有";
            }
        }
        return "";
    }

    /**
     * 返回文件状态
     * @param matching
     * @return
     */
    public static String getStorageState(String matching){
        if(!StringUtils.isEmpty(matching) && STORAGESTATE_ONLINE.equals(matching)){
            return "在线";
        }
        if(!StringUtils.isEmpty(matching) && STORAGESTATE_INTO_ONLINE.equals(matching)){
            return "进线";
        }
        if(!StringUtils.isEmpty(matching) && STORAGESTATE_OFFONLINE.equals(matching)){
            return "离线";
        }
        return "";
    }


    /**
     * 这里是将百纳秒转化成时间格式
     * @param nanosecond_100
     * @return
     */
    public static String getTimeLength(long nanosecond_100){
        String timeStr = l100ns2TC(nanosecond_100);
        return timeStr;
    }

    public static String FormatTimeCodeString(long hours, long minutes, long seconds, long frames, boolean dropFrame) {
        String framesSeparator = ":";
        if(dropFrame) {
            framesSeparator = ".";
        }

        return String.format("%1$02d:%2$02d:%3$02d%5$s%4$02d", new Object[]{Long.valueOf(hours), Long.valueOf(minutes), Long.valueOf(seconds), Long.valueOf(frames), framesSeparator});
    }

    public static String l100ns2TC(long l100ns) {
        long dHour = (long)Math.floor((double)l100ns / (3600.0D * Math.pow(10.0D, 7.0D)));
        long llResidue = (long)((double)l100ns % (3600.0D * Math.pow(10.0D, 7.0D)));
        long dMin = (long)Math.floor((double)llResidue / (60.0D * Math.pow(10.0D, 7.0D)));
        llResidue %= (long)Math.floor(60.0D * Math.pow(10.0D, 7.0D));
        long dSec = (long)Math.floor((double)llResidue / Math.pow(10.0D, 7.0D));
        llResidue %= (long)Math.pow(10.0D, 7.0D);
        long dFraction = (long)Math.floor((double)llResidue / 100000.0D);
        return FormatTimeCodeString(dHour, dMin, dSec, dFraction, true);
    }

}
