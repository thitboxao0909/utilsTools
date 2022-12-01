package org.example;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBException;
import javax.xml.bind.Marshaller;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;

public class Main {
    public static void main(String[] args) throws Exception {
        System.out.println("Hello world!");

        Customer cus = new Customer("Hoang", "Hanoi", "tbx@gmail.com");
        documentObjectOutput docOut = new documentObjectOutput();

        docOut.setName(cus.getName());
        docOut.setAddress(cus.getAddress());
        docOut.setEmail(cus.getEmail());
        docOut.setFileName("C:\\Users\\thitb\\Desktop\\Common\\export2docx\\src\\main\\java\\org\\example\\NameOut.doc");
        docOut.setFileTemplatePath("C:\\Users\\thitb\\Desktop\\Common\\export2docx\\src\\main\\java\\org\\example\\Name.docx");

        String fileTemplatePath = docOut.getFileTemplatePath();
        String fileName = docOut.getFileName();

        String xml = fwdObjectToXML(docOut);
        System.out.println("xml==============================" + xml);
        InputStream is = Files.newInputStream(Paths.get(fileTemplatePath));
        WordprocessingMLPackage wordPackage = Docx4J.load(is);
        is.close();

        Docx4J.bind(wordPackage, xml,
                Docx4J.FLAG_BIND_INSERT_XML | Docx4J.FLAG_BIND_BIND_XML | Docx4J.FLAG_BIND_REMOVE_SDT);

        ByteArrayOutputStream baosWord = new ByteArrayOutputStream();
        Docx4J.save(wordPackage, baosWord);
        byte[] bWord = baosWord.toByteArray();
        baosWord.close();

        Docx4J.save(wordPackage, new File(fileName), Docx4J.FLAG_NONE);
        System.out.println("Saved: " + fileName);
    }

    private static String fwdObjectToXML(documentObjectOutput model) throws Exception {
        String xmlContent = "";
        try {
            // Create JAXB Context
            JAXBContext jaxbContext = JAXBContext.newInstance(documentObjectOutput.class);

            // Create Marshaller
            Marshaller jaxbMarshaller = jaxbContext.createMarshaller();

            // Required formatting??
            jaxbMarshaller.setProperty(Marshaller.JAXB_FORMATTED_OUTPUT, Boolean.TRUE);

            // Print XML String to Console
            StringWriter sw = new StringWriter();

            // Write XML to StringWriter
            jaxbMarshaller.marshal(model, sw);

            // Verify XML Content
            xmlContent = sw.toString();

        } catch (JAXBException e) {
            e.printStackTrace();
        }

        return xmlContent;
    }
}