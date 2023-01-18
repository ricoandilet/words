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

import com.google.common.collect.Lists;
import com.spire.doc.Document;
import com.spire.doc.FieldType;
import com.spire.doc.FileFormat;
import com.spire.doc.HeaderFooter;
import com.spire.doc.Section;
import com.spire.doc.documents.BreakType;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.documents.XHTMLValidationType;
import com.spire.doc.fields.TextRange;
import com.youland.words.model.DocumentHtmlAndFooter;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.jsoup.helper.ValidationException;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.util.CollectionUtils;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.List;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

public class DocumentConvert {

  private static final ThreadPoolExecutor threadPoolExecutor = new ThreadPoolExecutor(15, 50, 10, TimeUnit.SECONDS, new LinkedBlockingQueue<>(20), new ThreadPoolExecutor.CallerRunsPolicy());

  /**
   * generate doc by html.
   *
   * @param html word html template
   * @return ByteArrayResource
   */
  public static ByteArrayResource generateWordByHtml(String html) {

    Document doc = new Document();
    doc.loadFromStream(new ByteArrayInputStream(html.getBytes()),
            FileFormat.Html,
            XHTMLValidationType.None);
    ByteArrayOutputStream out = new ByteArrayOutputStream();
    doc.saveToFile(out, FileFormat.Docx_2013);
    return removeLogo(new ByteArrayResource(out.toByteArray()));
  }

  /**
   * @param htmlAndFooters word html template
   * @return ByteArrayResource
   */
  public static ByteArrayResource generateWordByHtml(List<DocumentHtmlAndFooter> htmlAndFooters) throws Exception {

    if (CollectionUtils.isEmpty(htmlAndFooters)) {
      throw new ValidationException("No documents");
    }
    // first document
    CompletableFuture<ByteArrayResource> firstFuture = CompletableFuture.supplyAsync(
            () -> {
              DocumentHtmlAndFooter first = htmlAndFooters.get(0);
              ByteArrayOutputStream firstOut = generateWord(first);
              ByteArrayResource firstResource = removeLogo(new ByteArrayResource(firstOut.toByteArray()));
              return firstResource;
            }, threadPoolExecutor);
    // add other documents
    List<CompletableFuture<ByteArrayResource>> featureList = Lists.newArrayList(firstFuture);
    CompletableFuture[] cfArray = new CompletableFuture[htmlAndFooters.size()];
    for(int i=1; i<htmlAndFooters.size();i++){
      DocumentHtmlAndFooter item = htmlAndFooters.get(i);
      CompletableFuture<ByteArrayResource> future = CompletableFuture.supplyAsync(
              () -> {
                ByteArrayOutputStream pdf = generateWord(item);
                ByteArrayResource byteArrayResource = removeLogo(new ByteArrayResource(pdf.toByteArray()));
                return byteArrayResource;
              }, threadPoolExecutor);
      featureList.add(future);
    }
    CompletableFuture.allOf(featureList.toArray(cfArray)).join();

    // append documents
    Document document = new Document(firstFuture.get().getInputStream());
    ByteArrayOutputStream out = new ByteArrayOutputStream();
    for (int i=1; i<featureList.size();i++){
      InputStream append = featureList.get(i).get().getInputStream();
      document.insertTextFromStream(append, FileFormat.Docx_2013);
    }
    document.saveToFile(out, FileFormat.Docx_2013);
    return removeLogo(new ByteArrayResource(out.toByteArray()));
  }

  private static ByteArrayOutputStream generateWord(DocumentHtmlAndFooter docHtmlAndFooter) {

    String htmlContent = docHtmlAndFooter.getDocumentHtml();
    DocumentHtmlAndFooter.Footer docFooter = docHtmlAndFooter.getFooter();
    Document document = new Document();
    document.loadFromStream(new ByteArrayInputStream(htmlContent.getBytes()),
        FileFormat.Html,
        XHTMLValidationType.None);
    // add footer
    Section section = document.getSections().get(0);
    section.getPageSetup().setFooterDistance(15f);
    // get footer
    HeaderFooter footer = section.getHeadersFooters().getFooter();
    Paragraph footerParagraph = footer.addParagraph();
    section.getPageSetup().setRestartPageNumbering(true);
    section.getPageSetup().setPageStartingNumber(1);
    // set footer information
    TextRange first = footerParagraph.appendText(docFooter.getTitle().concat(" - Page "));
    TextRange second = footerParagraph.appendField("page number", FieldType.Field_Page);
    TextRange third = footerParagraph.appendText(" of ");
    TextRange fourth = footerParagraph.appendText(String.valueOf(document.getPageCount()));
    footerParagraph.appendBreak(BreakType.Line_Break);
    TextRange fifth = footerParagraph.appendText("Loan ID: ".concat(docFooter.getLoanId()));
    footerParagraph.appendBreak(BreakType.Line_Break);
    TextRange sixth = footerParagraph.appendText("Property Address: ".concat(docFooter.getAddress()));
    first.getCharacterFormat().setFontSize(10f);
    second.getCharacterFormat().setFontSize(10f);
    third.getCharacterFormat().setFontSize(10f);
    fourth.getCharacterFormat().setFontSize(10f);
    fifth.getCharacterFormat().setFontSize(10f);
    first.getCharacterFormat().setFontSize(10f);
    sixth.getCharacterFormat().setFontSize(10f);
    // set the location
    footerParagraph.getFormat().setLeftIndent(-20);

    ByteArrayOutputStream out = new ByteArrayOutputStream();
    document.saveToFile(out, FileFormat.Docx_2013);

    return out;
  }

  private static ByteArrayResource removeLogo(ByteArrayResource byteArrayResource) {
    try {
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
