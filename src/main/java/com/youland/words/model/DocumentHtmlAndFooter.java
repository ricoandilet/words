package com.youland.words.model;

import com.youland.words.core.Footer;
import com.youland.words.core.Margins;
import lombok.Builder;
import lombok.Data;

/**
 * @author: rico
 * @date: 2023/1/10
 **/
@Data
@Builder
public class DocumentHtmlAndFooter {

   private String documentHtml;

   private Footer footer;

   private Margins margins;

}
