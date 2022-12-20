package org.example;

import io.github.millij.poi.ss.model.annotations.SheetColumn;

public class Meo {

    @SheetColumn("Name")
    private String name;

    @SheetColumn("Age")
    private Integer Age;

    public Meo(String name, Integer age) {
        this.name = name;
        Age = age;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return Age;
    }

    public void setAge(Integer age) {
        Age = age;
    }
}
