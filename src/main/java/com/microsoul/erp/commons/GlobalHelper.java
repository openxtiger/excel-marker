package com.microsoul.erp.commons;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 微妞分布式平台-公用工具包
 *
 * @author 广州加叁信息科技有限公司 (tiger@microsoul.com)
 * @version V1.0.0
 */
public class GlobalHelper {
    private static String resHost;


    public static String toString(Object s) {
        return s == null ? "" : s.toString().trim();
    }


    public static long parseLong(String v, long defaultValue) {
        if (v == null || v.trim().equals(""))
            return defaultValue;
        try {
            return Long.parseLong(v);
        } catch (Exception e) {
            return defaultValue;
        }
    }


    public static int parseInteger(String v, int defaultValue) {
        if (v == null || v.trim().equals(""))
            return defaultValue;
        try {
            return Integer.parseInt(v);
        } catch (Exception e) {
            return defaultValue;
        }
    }

    public static int parseInt(Object v, int defaultValue) {
        if (v == null) return defaultValue;
        if (v instanceof Number) return (Integer) v;
        return parseInteger((String) v, defaultValue);
    }

    public static double parseDouble(String v, double defaultValue) {
        if (v == null || v.trim().equals(""))
            return defaultValue;
        try {
            return Double.parseDouble(v);
        } catch (Exception e) {
            return defaultValue;
        }
    }


    public static double round(double v, int scale) {
        try {
            if (Double.isInfinite(v)) return 0;
            BigDecimal b = new BigDecimal(Double.toString(v));
            BigDecimal one = new BigDecimal("1");
            return b.divide(one, scale, BigDecimal.ROUND_HALF_UP).doubleValue();
        } catch (Exception e) {
            System.err.println("err:" + v);
            throw e;
        }
    }


    public static ArrayList<String> split(String regex, String input) {
        Pattern pt = Pattern.compile(regex);
        Matcher mc = pt.matcher(input);
        int index = 0;
        ArrayList<String> matchList = new ArrayList<String>();
        while (mc.find()) {
            String match = input.subSequence(index, mc.start()).toString();
            matchList.add(match);
            matchList.add(mc.group(1));
            index = mc.end();
        }
        if (index <= input.length() - 1) {
            matchList.add(input.subSequence(index, input.length()).toString());
        }
        return matchList;
    }

    public static boolean isEmpty(Object str) {
        return str == null || "".equals(str);
    }

    public static String getExtension(String filename) {
        if (filename == null) {
            return null;
        } else {
            int index = indexOfExtension(filename);
            return index == -1 ? "" : filename.substring(index + 1);
        }
    }

    public static int indexOfExtension(String filename) {
        if (filename == null) {
            return -1;
        } else {
            int extensionPos = filename.lastIndexOf(46);
            int lastSeparator = indexOfLastSeparator(filename);
            return lastSeparator > extensionPos ? -1 : extensionPos;
        }
    }

    public static int indexOfLastSeparator(String filename) {
        if (filename == null) {
            return -1;
        } else {
            int lastUnixPos = filename.lastIndexOf(47);
            int lastWindowsPos = filename.lastIndexOf(92);
            return Math.max(lastUnixPos, lastWindowsPos);
        }
    }



    private static String padding(int repeat, char padChar) throws IndexOutOfBoundsException {
        if (repeat < 0) {
            throw new IndexOutOfBoundsException("Cannot pad a negative amount: " + repeat);
        } else {
            char[] buf = new char[repeat];

            for (int i = 0; i < buf.length; ++i) {
                buf[i] = padChar;
            }

            return new String(buf);
        }
    }


    public static String leftPad(String str, int size, char padChar) {
        if (str == null) {
            return null;
        } else {
            int pads = size - str.length();
            return pads <= 0 ? str : (pads > 8192 ? leftPad(str, size, String.valueOf(padChar)) : padding(pads, padChar).concat(str));
        }
    }

    public static String leftPad(String str, int size, String padStr) {
        if (str == null) {
            return null;
        } else {
            if (isEmpty(padStr)) {
                padStr = " ";
            }

            int padLen = padStr.length();
            int strLen = str.length();
            int pads = size - strLen;
            if (pads <= 0) {
                return str;
            } else if (padLen == 1 && pads <= 8192) {
                return leftPad(str, size, padStr.charAt(0));
            } else if (pads == padLen) {
                return padStr.concat(str);
            } else if (pads < padLen) {
                return padStr.substring(0, pads).concat(str);
            } else {
                char[] padding = new char[pads];
                char[] padChars = padStr.toCharArray();

                for (int i = 0; i < pads; ++i) {
                    padding[i] = padChars[i % padLen];
                }

                return (new String(padding)).concat(str);
            }
        }
    }

}
