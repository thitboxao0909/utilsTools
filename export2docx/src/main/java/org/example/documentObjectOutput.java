package org.example;

import java.io.Serializable;
import javax.xml.bind.annotation.XmlRootElement;

@XmlRootElement(name="documentObjectXML")
public class documentObjectOutput implements Serializable {
    private static final long serialVersionUID = 1L;
    private String fileTemplatePath;
    private String fileName;
    private String name;
    private String email;
    private String address;


    public String getFileTemplatePath() {
        return fileTemplatePath;
    }

    public void setFileTemplatePath(String fileTemplatePath) {
        this.fileTemplatePath = fileTemplatePath;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }
}
