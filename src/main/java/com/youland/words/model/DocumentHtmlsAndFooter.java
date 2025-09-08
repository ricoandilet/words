package com.youland.words.model;

import com.youland.words.core.Footer;
import com.youland.words.core.Margins;
import lombok.Builder;
import lombok.Data;

import java.util.List;

/**
 * @author: rico
 * @date: 2023/1/10
 **/
@Data
@Builder
public class DocumentHtmlsAndFooter {

   private List<String> documentHtmls;

   private Footer footer;

   private Margins margins;
}
