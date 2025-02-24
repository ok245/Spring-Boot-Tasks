package com.app.pojos;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import lombok.Data;

@Data
@Entity
public class Employee {

    @Id
    private Long id;
    private String firstName;
    private String lastName;
    private String email;
    private String department;
    private Double salary;

}
