package org.example;

import org.swagger2html.Swagger2Html;

import java.awt.*;
import java.io.*;

public class Main {
    public static void main(String[] args) throws Exception {
        System.out.println("Hello world!");

        Swagger2Html s2h = new Swagger2Html();
        String filePath = args[0];
        String jsonName = args[1];
        String htmlName = args[2];
        String jsonPath = filePath + "\\" + jsonName;
        String htmlPath = filePath + "\\" + htmlName;
        FileOutputStream fileOutputStream = new FileOutputStream(htmlPath);
        Writer writer = new java.io.OutputStreamWriter(fileOutputStream, "utf8");
        BufferedWriter bw = new BufferedWriter(writer);

        //Writer writer = new FileWriter(htmlPath,);

        s2h.toHtml(jsonPath, bw);

    }
}