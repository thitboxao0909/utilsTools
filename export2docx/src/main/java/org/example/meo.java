package org.example;

public class meo {
    public String catName;
    public int catAge;

    public meo(String catName, int catAge) {
        this.catName = catName;
        this.catAge = catAge;
    }

    @Override
    public String toString() {
        return "meo{" +
                "catName='" + catName + '\'' +
                ", catAge=" + catAge +
                '}';
    }
}
