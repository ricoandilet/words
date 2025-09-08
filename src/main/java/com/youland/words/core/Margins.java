package com.youland.words.core;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;

/**
 * @author: rico
 * @date: 2023/2/24
 */
@Data
@AllArgsConstructor
public class Margins {

  private float left = 72F;
  private float top = 72F;
  private float right = 72F;
  private float bottom = 72F;

}
