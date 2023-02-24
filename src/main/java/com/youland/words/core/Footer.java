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
package com.youland.words.core;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;

/**
 * @author: rico
 * @date: 2023/2/24
 */
@Data
@Builder
@AllArgsConstructor
public class Footer {

  private String title;

  private String LoanId;

  private String address;
}
