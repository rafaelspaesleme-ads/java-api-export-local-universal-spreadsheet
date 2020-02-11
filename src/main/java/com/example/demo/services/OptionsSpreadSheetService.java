package com.example.demo.services;

import java.io.IOException;
import java.util.List;

public class OptionsSpreadSheetService {

    /**
     * @title Exportador e conversor de planilhas
     *
     * @comment Api para exportar dados convertidos para uma planilha de Excel, nos seguintes formatos: XLS, XLSX, ODS.
     *
     * @code Tipos de Content
     * @code Type: XLS: application/vnd.ms-excel
     * @code XLSX: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
     * @code ODS: application/vnd.oasis.opendocument.spreadsheet
     *
     * @param dirName
     * @param fileName
     * @param stringContentList
     * @param stringHeaderList
     * @param titleTag
     * @param typeContent
     *
     * @throws IOException
     *
     * @since 11-02-2020
     *
     * @version 1.0
     *
     **/

    public static void exportSpreadSheet(List<String> stringHeaderList, List<List<String>> stringContentList, String dirName, String fileName, String titleTag, String typeContent, String lang, String country) throws IOException {
        switch (typeContent) {
            case "application/vnd.ms-excel":
                XlsService.exportExcel(stringHeaderList, stringContentList, dirName, fileName, titleTag);
                break;
            case "application/vnd.oasis.opendocument.spreadsheet":
                OdsService.exportExcel(stringHeaderList, stringContentList, dirName, fileName, titleTag, lang, country);
                break;
            case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                XlsxService.exportExcel(stringHeaderList, stringContentList, dirName, fileName, titleTag);
                break;
        }
    }

}
