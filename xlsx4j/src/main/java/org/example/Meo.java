package org.example;

public class Meo {

    private String name;

    private Integer Age;

    private String breeds;

    public Meo(String name, Integer age, String breeds) {
        this.name = name;
        this.Age = age;
        this.breeds = breeds;
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

    public String getBreeds() {
        return breeds;
    }

    public void setBreeds(String breeds) {
        this.breeds = breeds;
    }
}
