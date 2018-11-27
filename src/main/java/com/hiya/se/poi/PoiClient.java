package com.hiya.se.poi;

public class PoiClient
{
    public static void main(String[] args)
    {
        String wordPath = "E://poi/hiya.docx";
        IPio word =  IPio.create(HiyaPoiWord::new);
        word.doCreate(wordPath);
        word.doParse(wordPath);
        
        IPio excel =  IPio.create(HiyaPoiExcel::new);
        excel.doCreate(wordPath);
        excel.doParse(wordPath);
        
        IPio ppt =  IPio.create(HiyaPoiPPT::new);
        ppt.doCreate(wordPath);
        ppt.doParse(wordPath);
    }
}
