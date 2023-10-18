package org.example;

import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@NoArgsConstructor
public class Department {
    private String faculty;
    private List<Speciality> specialities;

    public Department(String faculty) {
        this.faculty = faculty;
    }
}
