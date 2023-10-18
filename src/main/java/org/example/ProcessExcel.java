package org.example;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * The {@code ProcessExcel} class is responsible for processing Excel workbooks and extracting relevant data.
 * It provides methods for extracting department name, speciality name, dean name, and processing the workbook to create a {@link Department} object.
 * Additionally, it contains utility methods for finding the first occurrence of a search string in a workbook and retrieving cell values.
 */
public class ProcessExcel {
    private static final String FACULTY = "Факультет";
    private static final String MONDAY = "Понеділок";
    private static final String SPECIALITY = "Спеціальність";
    private static final String DEAN = "Декан факультету";

    private String tempTime = null;
    private String tempWeekday = null;
    private String specialityRegex = "«(.*?)»";


    String extractDean(Workbook workbook) {
        Sheet sheet = workbook.getSheetAt(0);
        StringBuilder deanString = new StringBuilder();
        sheet.forEach(row -> row.forEach(cell -> {
            if (getCellValue(cell).contains(DEAN)) {
                deanString.append(getCellValue(cell).substring(getCellValue(cell).indexOf(DEAN)));
            }
        }));
        return deanString.toString().replace("_", "");
    }

    List<Department> processData(List<String> paths) {
        return paths.stream().map(path -> {
            File file = new File(path);
            try {
                Workbook workbook = WorkbookFactory.create(file);
                if (!extractDean(workbook).contains("Глущенко"))
                    return processIpzWorkbook(workbook);
                else
                    return processFenWorkbook(workbook);
            } catch (Exception e) {
                throw new RuntimeException("An error occurred while processing workbook", e);
            }
        }).toList();
    }

    private Department processIpzWorkbook(Workbook workbook) {
        String departmentName = extractDepartmentName(workbook);
        String specialityName = extractSpecialityName(workbook);
        int firstRow = findFirstOccurrenceOfSearchStringInWorkbook(workbook, MONDAY);
        List<Discipline> disciplines = extractDisciplines(workbook.getSheetAt(0), firstRow);
        return createDepartment(departmentName, specialityName, disciplines);
    }

    private Department processFenWorkbook(Workbook workbook) {
        String departmentName = extractDepartmentName(workbook);
        List<String> specialityNames = extractFenSpecialityName(workbook);
        int firstRow = findFirstOccurrenceOfSearchStringInWorkbook(workbook, MONDAY);
        List<Discipline> disciplines = extractDisciplines(workbook.getSheetAt(0), firstRow);
        return createFenDepartment(departmentName, specialityNames, disciplines);
    }

    private String extractDepartmentName(Workbook workbook) {
        Sheet sheet = workbook.getSheetAt(0);
        StringBuilder departmentString = new StringBuilder();
        sheet.forEach(row -> row.forEach(cell -> {
            if (getCellValue(cell).contains(FACULTY)) {
                departmentString.append(getCellValue(cell).substring(getCellValue(cell).indexOf(FACULTY)));
            }
        }));
        return departmentString.toString();
    }

    private String extractSpecialityName(Workbook workbook) {
        Sheet sheet = workbook.getSheetAt(0);
        StringBuilder specialityString = new StringBuilder();
        sheet.forEach(row -> row.forEach(cell -> {
            if (getCellValue(cell).contains(SPECIALITY)) {
                specialityString.append(getCellValue(cell).substring(getCellValue(cell).indexOf(SPECIALITY)));
            }
        }));
        return specialityString.toString().replace("\"", "");
    }

    private List<String> extractFenSpecialityName(Workbook workbook) {
        Sheet sheet = workbook.getSheetAt(0);
        Pattern pattern = Pattern.compile(specialityRegex);


        ArrayList<String> specialities = new ArrayList<>();
        sheet.forEach(row -> row.forEach(cell -> {
            if (getCellValue(cell).contains(SPECIALITY)) {
                Matcher matcher = pattern.matcher(getCellValue(cell));
                while (matcher.find()) {
                    specialities.add(matcher.group(1));
                }
            }
        }));
        return specialities;
    }

    private List<Discipline> extractDisciplines(Sheet sheet, int firstRow) {
        List<Discipline> disciplines = new ArrayList<>();
        int lastRow = sheet.getLastRowNum();
        for (int i = firstRow; i <= lastRow; i++) {
            Row row = sheet.getRow(i);
            Discipline discipline = processRow(row);
            if (discipline != null) {
                disciplines.add(discipline);
            }
        }
        return disciplines;
    }

    private Department createDepartment(String departmentName, String specialityName, List<Discipline> disciplines) {
        Department department = new Department(departmentName);
        Speciality speciality = new Speciality(specialityName);
        speciality.setDisciplines(disciplines);
        department.setSpecialities(List.of(speciality));
        return department;
    }

    private Department createFenDepartment(String departmentName, List<String> specialityNames, List<Discipline> disciplines) {
        Department department = new Department(departmentName);
        List<Speciality> specialities = specialityNames.stream().map(Speciality::new).toList();
        specialities.forEach(i -> i.setDisciplines(new ArrayList<>()));
        department.setSpecialities(specialities);
        department = insertSpecialities(department, disciplines);
        return department;
    }

    private Department insertSpecialities(Department department, List<Discipline> disciplines) {
        disciplines.forEach(i -> {
            String regex = "\\(([^)]+)\\)";
            String regex2 = "^[^()]*$";
            Pattern pattern = Pattern.compile(regex);
            Matcher matcher = pattern.matcher(i.getName());
            Pattern pattern2 = Pattern.compile(regex2);
            Matcher matcher2 = pattern2.matcher(i.getName());
            if(matcher2.find()){
                department.getSpecialities().forEach(speciality -> speciality.getDisciplines().add(filterFenDiscipline(i)));
            }
            while (matcher.find()) {
                String match = matcher.group(1).trim();
                String[] splitSpecialities;

                final boolean[] patternFound = {false};

                if (match.contains("+")) {
                    splitSpecialities = match.split("\\+");
                    patternFound[0] = true;
                } else if (match.contains(".")) {
                    splitSpecialities = match.split("\\.");
                    patternFound[0] = true;
                } else if (match.contains(",")) {
                    splitSpecialities = match.split(",");
                    patternFound[0] = true;
                } else {
                    splitSpecialities = new String[]{match};
                    patternFound[0] = true;
                }

                department.getSpecialities().forEach(speciality -> {
                    for (String name : splitSpecialities) {
                        if (speciality.getName().toLowerCase().contains(name.toLowerCase())) {
                            speciality.getDisciplines().add(filterFenDiscipline(i));
                            patternFound[0] = true;
                        }
                    }
                });
            }


        });
        return department;
    }

    public Discipline filterFenDiscipline(Discipline discipline) {
        if (discipline.getGroup().toLowerCase().contains("лекція")) discipline.setGroup("Лекція");
        if (discipline.getAuditorium().equalsIgnoreCase("Д")) discipline.setAuditorium("Дистанційно");

        String regex = "\\d+";

        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(discipline.getGroup());

        if (matcher.find()) {
            discipline.setGroup(matcher.group());
        }
        return discipline;
    }

    public Discipline processRow(Row row) {
        Iterator<Cell> cellIterator = row.cellIterator();
        String day = processDay(cellIterator);
        String time = processTime(cellIterator);
        String discipline = processCellValue(cellIterator);
        String group = processCellValue(cellIterator);
        String weeks = processCellValue(cellIterator);
        String auditorium = processCellValue(cellIterator);

        if (discipline == null || discipline.isEmpty()) {
            return null;
        }

        return new Discipline(discipline, day, time, group, weeks, auditorium);
    }

    private String processDay(Iterator<Cell> cellIterator) {
        String day = tempWeekday;
        if (cellIterator.hasNext()) {
            day = getCellValue(cellIterator.next());
            if (day.isEmpty()) {
                day = tempWeekday;
            } else {
                tempWeekday = day;
            }
        }
        return day;
    }

    private String processTime(Iterator<Cell> cellIterator) {
        String time = tempTime;
        if (cellIterator.hasNext()) {
            time = getCellValue(cellIterator.next());
            if (time.isEmpty()) {
                time = tempTime;
            } else {
                tempTime = time;
            }
        }
        return time;
    }

    private String processCellValue(Iterator<Cell> cellIterator) {
        if (cellIterator.hasNext()) {
            return getCellValue(cellIterator.next());
        }
        return null;
    }

    String getCellValue(Cell cell) {
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf((int) cell.getNumericCellValue());
            default -> "";
        };
    }

    /**
     * Helper method to retrieve the trimmed string value of a cell
     * @param cell The cell for which to retrieve the trimmed string
     * @return The trimmed string value of the cell
     */
    private String getTrimmedCellValue(Cell cell) {
        return cell.getStringCellValue().trim();
    }

    /**
     * This method finds the first occurrence of a search string in a Workbook,
     * searching row by row, cell by cell.
     * @param workbook The Workbook object to search
     * @param searchString The string to find in the Workbook
     * @return Returns the row number where the searchString was first found, or -1 if not found
     */
    int findFirstOccurrenceOfSearchStringInWorkbook(Workbook workbook, String searchString) {
        Sheet firstSheet = workbook.getSheetAt(0);
        for (Row currentRow : firstSheet) {
            for (Cell currentCell : currentRow) {
                if (searchString.equals(getTrimmedCellValue(currentCell))) {
                    return currentRow.getRowNum();
                }
            }
        }
        return -1;
    }
}