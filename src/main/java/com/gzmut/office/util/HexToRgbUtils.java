package com.gzmut.office.util;

import java.util.ArrayList;
import java.util.List;

/**
 * @author livejq
 * @since 2019/7/29
 */
public class HexToRgbUtils {

    public static String toRgb(String hex) {
        List<String> strs = splitToDouble(hex);
        String[] cha = new String[]{"A", "B", "C", "D", "E", "F"};
        Integer[] result = new Integer[6];
        if (strs != null) {
            for (int i = 0; i < strs.size(); i++) {
                for (int k = 0; k < cha.length; k++) {
                    if (strs.get(i).equals(cha[k])) {
                        if (i % 2 == 0) {
                            result[i] = (k + 10) * 16;
                        } else {
                            result[i] = k + 10;
                        }
                        break;
                    }
                }
                if (result[i] == null) {
                    if (i % 2 == 0) {
                        result[i] = Integer.parseInt(strs.get(i)) * 16;
                    } else {
                        result[i] = Integer.parseInt(strs.get(i));
                    }
                }
            }
        }
        StringBuilder sb = new StringBuilder();
        for (int j = 0; j < result.length; j = j + 2) {
            int x = result[j] + result[j + 1];
            if (j < 4) {
                sb.append(String.valueOf(x)).append(",");
            } else {
                sb.append(String.valueOf(x));
            }
        }

        return sb.toString();
    }

    private static List<String> splitToDouble(String hex) {
        if (hex.length() != 6) {
            System.out.println("请输入长度为6位十六进制颜色的字符串：如 00FF02");
            return null;
        }
        String index1 = hex.substring(0, 1);
        String index2 = hex.substring(1, 2);
        String index3 = hex.substring(2, 3);
        String index4 = hex.substring(3, 4);
        String index5 = hex.substring(4, 5);
        String index6 = hex.substring(5, 6);
        List<String> list = new ArrayList<>(6);
        list.add(index1);
        list.add(index2);
        list.add(index3);
        list.add(index4);
        list.add(index5);
        list.add(index6);

        return list;
    }
}
