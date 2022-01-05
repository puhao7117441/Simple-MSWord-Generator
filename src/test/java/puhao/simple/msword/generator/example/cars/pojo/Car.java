package puhao.simple.msword.generator.example.cars.pojo;

import java.time.LocalDate;

public class Car {
    private String model;
    private String price;
    private LocalDate boughtAt;

    public Car(String model, String price, LocalDate boughtAt){
        this.model = model;
        this.price = price;
        this.boughtAt = boughtAt;
    }

    public String getModel() {
        return model;
    }

    public void setModel(String model) {
        this.model = model;
    }

    public String getPrice() {
        return price;
    }

    public void setPrice(String price) {
        this.price = price;
    }

    public LocalDate getBoughtAt() {
        return boughtAt;
    }

    public void setBoughtAt(LocalDate boughtAt) {
        this.boughtAt = boughtAt;
    }
}
