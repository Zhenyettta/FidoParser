package org.example;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.util.Arrays;
import java.util.List;

public class Main {
    public static void main(String[] args) throws Exception {
        File folder = new File("src/main/java/org/example/schedules");
        ProcessExcel processExcel = new ProcessExcel();
        Gson gson = new GsonBuilder().setPrettyPrinting().disableHtmlEscaping().create();
        String json = gson.toJson(processExcel.processData(folder));
        try (OutputStreamWriter writer = new OutputStreamWriter(new FileOutputStream("normalized_schedule.json"), StandardCharsets.UTF_8)) {
            writer.write(json);
        }
    }
}