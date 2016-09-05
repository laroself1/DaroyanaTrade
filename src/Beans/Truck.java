package Beans;

import  java.util.Date;
/**
 * Created by LaroSelf on 20.05.2016.
 */
public class Truck {
    private String position;
    private String date;
    private String invoice;
    private String number;
    private String weight;
    private String consignor;
    private String goodClass;
    private String warehouse;

    public Truck(String position, String date, String invoice, String number, String weight, String consignor, String goodClass, String warehouse) {
        this.position = position;
        this.date = date;
        this.invoice = invoice;
        this.number = number;
        this.weight = weight;
        this.consignor = consignor;
        this.goodClass = goodClass;
        this.warehouse = warehouse;
    }


    public String getPosition() {
        return position;
    }

    public void setPosition(String position) {
        this.position = position;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public String getInvoice() {
        return invoice;
    }

    public void setInvoice(String invoice) {
        this.invoice = invoice;
    }

    public String getNumber() {
        return number;
    }

    public void setNumber(String number) {
        this.number = number;
    }

    public String getWeight() {
        return weight;
    }

    public void setWeight(String weight) {
        this.weight = weight;
    }

    public String getConsignor() {
        return consignor;
    }

    public void setConsignor(String consignor) {
        this.consignor = consignor;
    }

    public String getGoodClass() {
        return goodClass;
    }

    public void setGoodClass(String goodClass) {
        this.goodClass = goodClass;
    }

    public String getWarehouse() {
        return warehouse;
    }

    public void setWarehouse(String warehouse) {
        this.warehouse = warehouse;
    }
}
