package com.youland.words.core;

import lombok.Getter;
import lombok.Setter;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author: rico
 * @date: 2025/3/7
 **/
@Setter
@Getter
public class ReplaceContent {

    private Map<String,String> textContentList = new HashMap<>();
    private Map<String,Boolean> checkBoxContentList = new HashMap<>();
    private Map<String,String> textBoxContentList = new HashMap<>();
}
