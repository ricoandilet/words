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

import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.documents.XHTMLValidationType;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.springframework.core.io.ByteArrayResource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;

public class WordUtil {

  /**
   * generate doc by html.
   *
   * @param html
   * @return ByteArrayResource
   */
  public static ByteArrayResource generateWord(String html) {

    Document doc = new Document();
    doc.loadFromStream(
        new ByteArrayInputStream(html.getBytes()),
        FileFormat.Html,
        XHTMLValidationType.None);
    ByteArrayOutputStream out = new ByteArrayOutputStream();
    doc.saveToFile(out, FileFormat.Docx_2013);

    return removeLogo(out);
  }

  public static ByteArrayResource generateWord(List<String> listHtml){

    return null;
  }

  private static ByteArrayResource removeLogo(ByteArrayOutputStream docResource) {

    try {

      ByteArrayResource byteArrayResource = new ByteArrayResource(docResource.toByteArray());
      ByteArrayOutputStream out = new ByteArrayOutputStream();

      XWPFDocument doc = new XWPFDocument(OPCPackage.open(byteArrayResource.getInputStream()));
      List<XWPFParagraph> paragraphs = doc.getParagraphs();

      if (paragraphs.size() < 1) return byteArrayResource;
      XWPFParagraph firstParagraph = paragraphs.get(0);
      if (firstParagraph.getText().contains("Spire.Doc")) {
        doc.removeBodyElement(doc.getPosOfParagraph(firstParagraph));
      }
      doc.write(out);
      return new ByteArrayResource(out.toByteArray());
    } catch (Exception e) {
      e.printStackTrace();
    }
    return null;
  }
}
