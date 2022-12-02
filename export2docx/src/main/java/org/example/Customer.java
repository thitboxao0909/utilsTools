package org.example;

import org.docx4j.wml.Style;

import java.util.List;

public class Customer {
    private String name;
    private String address;
    private String email;

    private List<meo> meo;

    public List<meo> getLstMeo() {
        return meo;
    }

    public void setLstMeo(List<meo> meo) {
        this.meo = meo;
    }

    public Customer(String name, String address, String email) {
        this.name = name;
        this.address = address;
        this.email = email;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }
}
