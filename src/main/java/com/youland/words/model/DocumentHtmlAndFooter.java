package com.youland.words.model;

import lombok.AllArgsConstructor;
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

   @Data
   @Builder
   @AllArgsConstructor
   public static class Footer{

      private String title;

      private String LoanId;

      private String address;
   }

}
