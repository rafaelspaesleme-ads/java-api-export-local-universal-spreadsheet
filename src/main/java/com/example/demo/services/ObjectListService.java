package com.example.demo.services;

import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;

public class ObjectListService {

    public static List<String> listHeader = new ArrayList<>();

    protected static List<String> builderHeaders(List<String> stringHeaderList){
        listHeader.addAll(stringHeaderList);
        return listHeader;
    }
}