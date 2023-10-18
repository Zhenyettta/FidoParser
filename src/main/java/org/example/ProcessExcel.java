package org.example;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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


    /**
     * Extracts a string containing the keyword "DEAN" from the given workbook.
     *
     * @param workbook the workbook from which to extract the dean string
     * @return the extracted dean string with underscores removed
     */
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

    /**
     * Processes a list of file paths and returns a list of Department objects.
     *
     * @param paths the list of file paths to process
     * @return a list of Department objects
     * @throws RuntimeException if an error occurs while processing the workbooks
     */
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

    /**
     * Processes an IPZ workbook and returns a Department object.
     *
     * @param workbook the IPZ workbook to process
     * @return a Department object
     */
    private Department processIpzWorkbook(Workbook workbook) {
        String departmentName = extractDepartmentName(workbook);
        String specialityName = extractSpecialityName(workbook);
        int firstRow = findFirstOccurrenceOfSearchStringInWorkbook(workbook, MONDAY);
        List<Discipline> disciplines = extractDisciplines(workbook.getSheetAt(0), firstRow);
        return createDepartment(departmentName, specialityName, disciplines);
    }

    /**
     * Processes the given workbook to create a department for the FEN (Faculty of Engineering) department.
     *
     * @param workbook The workbook to process.
     * @return The created department.
     */
    private Department processFenWorkbook(Workbook workbook) {
        String departmentName = extractDepartmentName(workbook);
        List<String> specialityNames = extractFenSpecialityName(workbook);
        int firstRow = findFirstOccurrenceOfSearchStringInWorkbook(workbook, MONDAY);
        List<Discipline> disciplines = extractDisciplines(workbook.getSheetAt(0), firstRow);
        return createFenDepartment(departmentName, specialityNames, disciplines);
    }

    /**
     * Extracts the department name from the given workbook.
     *
     * @param workbook The workbook to extract the department name from.
     * @return The extracted department name.
     */
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

    /**
     * Extracts the speciality name from the given workbook's sheet at index 0.
     *
     * @param workbook the workbook object containing the sheet with speciality name
     * @return the extracted speciality name as a String
     */
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

    /**
     * Extracts the FEN speciality names from the given workbook's sheet at index 0.
     *
     * @param workbook the workbook object containing the sheet with FEN speciality names
     * @return a list of extracted FEN speciality names as strings
     */
    private List<String> extractFenSpecialityName(Workbook workbook) {
        Sheet sheet = workbook.getSheetAt(0);
        String specialityRegex = "«(.*?)»";
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

    /**
     * Extracts a list of disciplines from the given sheet, starting from the specified first row.
     *
     * @param sheet    The sheet from which to extract the disciplines.
     * @param firstRow The index of
     */
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

    /**
     * Creates a new department with the given department name, speciality name, and list of disciplines.
     *
     * @param departmentName the name of the department
     * @param specialityName the name of the speciality
     * @param disciplines    the list of disciplines
     * @return the created department
     */
    private Department createDepartment(String departmentName, String specialityName, List<Discipline> disciplines) {
        Department department = new Department(departmentName);
        Speciality speciality = new Speciality(specialityName);
        speciality.setDisciplines(disciplines);
        department.setSpecialities(List.of(speciality));
        return department;
    }

    /**
     * Creates a Department object for the FEN department.
     *
     * @param departmentName  the name of the department
     * @param specialityNames the list of speciality names
     * @param disciplines     the list of disciplines
     * @return the Department object for the FEN department
     */
    private Department createFenDepartment(String departmentName, List<String> specialityNames, List<Discipline> disciplines) {
        Department department = new Department(departmentName);
        List<Speciality> specialities = specialityNames.stream().map(Speciality::new).toList();
        specialities.forEach(i -> i.setDisciplines(new ArrayList<>()));
        department.setSpecialities(specialities);
        department = insertSpecialities(department, disciplines);
        return department;
    }

    /**
     * Inserts disciplines into appropriate specialities in the given department.
     *
     * @param department  the department to insert specialities into
     * @param disciplines the list of disciplines to insert
     * @return the department with inserted specialities
     */
    private Department insertSpecialities(Department department, List<Discipline> disciplines) {
        final String parenthesesRegex = "\\(([^)]+)\\)";
        final String noParenthesesRegex = "^[^()]*$";
        final Pattern parenthesesPattern = Pattern.compile(parenthesesRegex);
        final Pattern noParenthesesPattern = Pattern.compile(noParenthesesRegex);

        for (Discipline discipline : disciplines) {
            Matcher matcher = parenthesesPattern.matcher(discipline.getName());
            Matcher matcher2 = noParenthesesPattern.matcher(discipline.getName());
            if (matcher2.find()) {
                addDisciplineToAllSpecialities(department, discipline);
            }

            while (matcher.find()) {
                String[] splitSpecialities = splitMatches(matcher.group(1).trim());
                addDisciplineToAppropriateSpecialities(department, discipline, splitSpecialities);
            }
        }

        return department;
    }

    /**
     * Adds a discipline to all specialities in the given department
     */
    private void addDisciplineToAllSpecialities(Department department, Discipline discipline) {
        department.getSpecialities().forEach(speciality -> speciality.getDisciplines().add(filterFenDiscipline(discipline)));
    }

    /**
     * Splits the given match string into an array of substrings based on certain delimiters.
     * If the match contains the delimiter '+', it will split the match using '+'.
     * If the match contains the delimiter '.', it will split the match using '.'.
     * If the match contains the delimiter ',', it will split the match using ','.
     * If the match does not contain any of the delimiters, it will create a new array with the match as the only element.
     *
     * @param match the string to be splitted
     * @return an array of
     */
    private String[] splitMatches(String match) {
        String[] splitSpecialities;

        if (match.contains("+")) {
            splitSpecialities = match.split("\\+");
        } else if (match.contains(".")) {
            splitSpecialities = match.split("\\.");
        } else if (match.contains(",")) {
            splitSpecialities = match.split(",");
        } else {
            splitSpecialities = new String[]{match};
        }

        return splitSpecialities;
    }

    /**
     * Adds a discipline to appropriate specialities in a department.
     *
     * @param department        the department to add the discipline to
     * @param discipline        the discipline to add
     * @param splitSpecialities an array of speciality names to match and add the discipline to
     */
    private void addDisciplineToAppropriateSpecialities(Department department, Discipline discipline, String[] splitSpecialities) {
        department.getSpecialities().forEach(speciality -> {
            for (String name : splitSpecialities) {
                if (speciality.getName().toLowerCase().contains(name.toLowerCase())) {
                    speciality.getDisciplines().add(filterFenDiscipline(discipline));
                }
            }
        });
    }

    /**
     * Filters the given discipline by modifying its group and auditorium properties.
     * If the discipline's group contains the substring "лекція" (case-insensitive), it will be modified to "Лекція".
     * If the discipline's auditorium is "Д" (case-insensitive), it will be modified to "Дистанційно".
     * The group will be modified to contain only digits.
     *
     * @param discipline The discipline to filter.
     * @return The filtered discipline.
     */
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

    /**
     * Processes a single row to extract information and create a Discipline object.
     *
     * @param row the row to process
     * @return a Discipline object created from the row information
     */
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

    /**
     * Processes the day value from the given cellIterator.
     *
     * @param cellIterator the iterator containing the cells to process
     * @return the processed day value
     */
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

    /**
     * Processes the time value from the given iterator.
     *
     * @param cellIterator the iterator of cells containing the time value
     * @return the processed time value
     */
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

    /**
     * Processes the value of a cell in the given iterator.
     *
     * @param cellIterator The iterator containing the cells.
     */
    private String processCellValue(Iterator<Cell> cellIterator) {
        if (cellIterator.hasNext()) {
            return getCellValue(cellIterator.next());
        }
        return null;
    }

    /**
     * Returns the value of a Cell as a String.
     *
     * @param cell The Cell to get the value from.
     * @return The value of the Cell as a String. If the Cell is of type STRING, the String value is returned.
     * If the Cell is of type NUMERIC, the
     */
    String getCellValue(Cell cell) {
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf((int) cell.getNumericCellValue());
            default -> "";
        };
    }

    /**
     * Helper method to retrieve the trimmed string value of a cell
     */
    private String getTrimmedCellValue(Cell cell) {
        return cell.getStringCellValue().trim();
    }

    /**
     * This method finds the first occurrence of a search string in a Workbook,
     * searching row by row, cell by cell.
     *
     * @param workbook     The Workbook object to search
     * @param searchString The string to find in the Workbook
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