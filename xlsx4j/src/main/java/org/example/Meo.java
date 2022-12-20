package org.example;

public class Meo {

    private String name;

    private Integer Age;


    public Meo(String name, Integer age) {
        this.name = name;
        this.Age = age;
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
