package org.example;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.example.Discipline;

import java.util.List;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class Speciality {
    private String name;
    private List<Discipline> disciplines;

    public Speciality(String name){
        this.name = name;
    }
}