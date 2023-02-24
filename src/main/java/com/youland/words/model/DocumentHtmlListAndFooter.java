package com.youland.words.model;

import com.youland.words.core.Footer;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;

import java.util.List;

/**
 * @author: rico
 * @date: 2023/1/10
 **/
@Data
@Builder
public class DocumentHtmlListAndFooter {

   private List<String> documentHtmlList;

   private Footer footer;
}
