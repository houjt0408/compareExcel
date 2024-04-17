package org.example;

public class util {
    public static String removeAfterLastDot(String str) {
        int lastDotIndex = str.lastIndexOf('.');
        if (lastDotIndex != -1) {
            return str.substring(0, lastDotIndex);
        } else {
            return str;
        }
    }
}
