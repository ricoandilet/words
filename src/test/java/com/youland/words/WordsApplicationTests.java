package com.youland.words;

import com.youland.words.core.DocumentConvert;
import com.youland.words.core.ReplaceContent;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.HashMap;
import java.util.Map;

class WordsApplicationTests {

    @Test
    void contextLoads() throws FileNotFoundException {
        ReplaceContent replaceContent = new ReplaceContent();
        Map<String,Boolean> checkBoxContentList = new HashMap<>();
        checkBoxContentList.put("checkbox_1", true);
        checkBoxContentList.put("checkbox_2", true);
        replaceContent.setCheckBoxContentList(checkBoxContentList);
        String filePath = "./src/test/data/fw9_v3.docx";
        FileInputStream fis = new FileInputStream(filePath);
        DocumentConvert.replaceDocument(replaceContent, fis);

    }

}
