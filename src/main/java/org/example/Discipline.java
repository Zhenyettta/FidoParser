package org.example;

import lombok.AllArgsConstructor;
import lombok.Data;

@Data
@AllArgsConstructor
public class Discipline {
    private String name;
    private String day;
    private String time;
    private String group;
    private String weeks;
    private String auditorium;
}