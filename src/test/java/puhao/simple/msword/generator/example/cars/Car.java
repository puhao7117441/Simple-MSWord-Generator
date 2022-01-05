package puhao.simple.msword.generator.example.cars;

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


    public LocalDate getBoughtAt() {
        return boughtAt;
    }


}
