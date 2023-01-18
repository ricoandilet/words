/*
  Copyright (C) 2018-2021 YouYu information technology (Shanghai) Co., Ltd.
  <p>
  All right reserved.
  <p>
  This software is the confidential and proprietary
  information of YouYu Company of China.
  ("Confidential Information"). You shall not disclose
  such Confidential Information and shall use it only
  in accordance with the terms of the contract agreement
  you entered into with YouYu inc.
 */
package com.youland.words.utils;

import java.util.Properties;

/**
 * @author: rico
 * @date: 2023/1/14
 **/
public class SystemUtil {

    private final static String mac = "Mac";
    private final static String window = "Windows";

    public static String getSystemName(){
        Properties properties = System.getProperties();
        return properties.getProperty("os.name");
    }

    public static boolean isLinux(){
        Properties properties = System.getProperties();
        String osName = properties.getProperty("os.name");
        if(!osName.contains(mac) && !osName.contains(window)){
            return true;
        }
        return false;
    }

  public static void main(String[] args) {
    System.out.println("SystemUtil.main: = "+SystemUtil.getSystemName());
  }
}
