package org.lyuban.apachepoi.model;

import jakarta.persistence.*;
import lombok.Setter;

import java.util.Objects;

@Entity
public class MaterialPosition {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;
    @Column(nullable = false)
    private String naimenovanie;
    private String measurementUnit;
    @Setter
    private double unitPrice;
    private double kgPrice;

    public MaterialPosition() {
    }

    public MaterialPosition(Long id, String naimenovanie, String measurementUnit, double unitPrice, double kgPrice) {
        this.id = id;
        this.naimenovanie = naimenovanie;
        this.measurementUnit = measurementUnit;
        this.unitPrice = unitPrice;
        this.kgPrice = kgPrice;
    }

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getNaimenovanie() {
        return naimenovanie;
    }

    public void setNaimenovanie(String naimenovanie) {
        this.naimenovanie = naimenovanie;
    }

    public String getMeasurementUnit() {
        return measurementUnit;
    }

    public void setMeasurementUnit(String measurementUnit) {
        this.measurementUnit = measurementUnit;
    }

    public double getUnitPrice() {
        return unitPrice;
    }

    public void setUnitPrice(double unitPrice) {
        this.unitPrice = unitPrice;
    }

    public double getKgPrice() {
        return kgPrice;
    }

    public void setKgPrice(double kgPrice) {
        this.kgPrice = kgPrice;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        MaterialPosition that = (MaterialPosition) o;
        return Double.compare(unitPrice, that.unitPrice) == 0 && Double.compare(kgPrice, that.kgPrice) == 0 && Objects.equals(id, that.id) && Objects.equals(naimenovanie, that.naimenovanie) && Objects.equals(measurementUnit, that.measurementUnit);
    }

    @Override
    public int hashCode() {
        return Objects.hash(id, naimenovanie, measurementUnit, unitPrice, kgPrice);
    }

    @Override
    public String toString() {
        return "MaterialPosition{" +
                "id=" + id +
                ", naimenovanie='" + naimenovanie + '\'' +
                ", measurementUnit='" + measurementUnit + '\'' +
                ", unitPrice=" + unitPrice +
                ", kgPrice=" + kgPrice +
                '}';
    }
}
