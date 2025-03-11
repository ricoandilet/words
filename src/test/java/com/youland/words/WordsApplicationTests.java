package com.youland.words;

import com.youland.words.core.DocumentConvert;
import com.youland.words.core.ReplaceContent;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.core.io.ByteArrayResource;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.HashMap;
import java.util.Map;

@Disabled
class WordsApplicationTests {

    @Test
    void contextLoads() throws FileNotFoundException {
        ReplaceContent replaceContent = new ReplaceContent();
        Map<String,Boolean> checkBoxContentList = new HashMap<>();
        checkBoxContentList.put("checkbox_1", true);
        checkBoxContentList.put("checkbox_2", true);

        Map<String,String> textBoxContent = new HashMap<>();
        textBoxContent.put("ssn_1", " ");
        textBoxContent.put("ssn_2", " ");
        textBoxContent.put("ssn_3", " ");
        textBoxContent.put("ssn_4", " ");
        textBoxContent.put("ssn_5", " ");
        textBoxContent.put("ssn_6", " ");
        textBoxContent.put("ssn_7", " ");
        textBoxContent.put("ssn_8", " ");
        textBoxContent.put("ssn_9", "1");
        replaceContent.setTextBoxContentList(textBoxContent);
        String filePath = "./src/test/data/fw9_v31.docx";
        FileInputStream fis = new FileInputStream(filePath);
        ByteArrayResource byteArrayResource = DocumentConvert.replaceDocument(replaceContent, fis);

    }

}
